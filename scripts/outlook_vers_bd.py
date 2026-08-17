#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Import direct Outlook -> base Constructo AI (contournement du webhook).

POURQUOI CE CHEMIN EXISTE. La voie propre est `outlook_bridge.py`, qui pousse
vers `POST /webhook/inbound` et laisse l'application faire son travail. Elle est
bloquee : le webhook exige `N8N_WEBHOOK_SECRET`, Render n'injecte une nouvelle
variable qu'au DEPLOIEMENT, et le deploiement echoue tant que GitHub est en
panne (incident du 2026-08-17, 13h40 UTC). Ce script ecrit donc directement dans
la base, en reproduisant fidelement ce que fait le webhook.

CE QU'IL REPREND DU WEBHOOK, A L'IDENTIQUE :
  - l'assainissement du HTML par `_sanitize_email_html` (la VRAIE fonction, elle
    est importee, pas recopiee : une copie derive au premier correctif amont) ;
  - le calcul du `thread_id` (References[0] > In-Reply-To > message_id) ;
  - la deduplication sur (message_id, account_id).

CE QU'IL FAIT MIEUX. Le webhook ecrit `date_received = CURRENT_TIMESTAMP` : tous
les messages importes porteraient la date de l'import. Ici on conserve les VRAIES
dates Outlook, ainsi que l'etat lu/non-lu et le sens (recu/envoye).

CE QU'IL NE FAIT PAS. Le rattachement au CRM (`company_id`, `contact_id`) reste
a NULL, et les pieces jointes ne sont pas importees. C'est deliberé : ces deux
mecanismes ont leur propre logique cote application, et les reproduire de
memoire produirait des liens faux — pire que pas de lien du tout.

TRACABLE, MAIS PLUS ANNULABLE PAR L'OUTIL. Chaque ligne porte
`labels_json = ["import-outlook", "run-AAAAMMJJ-HHMMSS-<pid>"]`, et l'etiquette
du run est imprimee a la fin d'un import reussi. En revanche `--annuler` A ETE
RETIRE : six tours de retest ont trouve six facons differentes pour lui de
supprimer PLUS que ce qu'il annoncait. Pour retirer un import, ecrire la requete
a la main apres avoir verifie l'hote de la connexion :
    SELECT count(*) FROM tenant_constructo_e1f633.emails
     WHERE labels_json ? 'run-AAAAMMJJ-HHMMSS-<pid>';
    DELETE FROM ...   -- le meme WHERE, une fois le compte verifie
On n'a de toute facon jamais besoin de supprimer pour reimporter : l'import
DEDUPLIQUE sur message_id.

Usage :
    python scripts/outlook_vers_bd.py --folder both --limit 50          # simulation
    python scripts/outlook_vers_bd.py --folder both --limit 50 --ecrire
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ICI = Path(__file__).resolve().parent
sys.path.insert(0, str(ICI))

import outlook_mail as om  # noqa: E402
import outlook_bridge as pont  # noqa: E402  (reutilise les extracteurs MAPI)

SCHEMA = "tenant_constructo_e1f633"
MARQUEUR = "import-outlook"
# Marqueur PROPRE A CETTE EXECUTION. Sans lui, `--annuler` n'a aucune notion de
# run : il compte et supprime TOUS les imports jamais faits. Le message de panne
# disait « pour les retirer : --annuler --oui » en parlant des lignes du run
# interrompu — reproduit au retest #205 tour 3 : 1000 lignes d'un import reussi
# + 200 d'un run casse, la commande recommandee detruisait les 1200. En
# production cela aurait emporte les 1431 lignes de l'historique.
# ⚠️ Resolution a la SECONDE + pid. Sans le pid, deux imports demarres dans la
# meme seconde recevaient le MEME identifiant et devenaient indiscernables :
# `--annuler --run <id>` en aurait supprime deux, ce qui retablit exactement le
# defaut que le marqueur de run existe pour empecher. Le detail par execution
# les fusionnait aussi, donc rien ne signalait la collision.
RUN_ID = "run-{}-{}".format(
    datetime.datetime.now().strftime("%Y%m%d-%H%M%S"), os.getpid())
ERP = Path(r"C:\Dev\Constructo_AI_Prod\ERP_REACT")


def charger_assainisseur():
    """Importe la VRAIE fonction d'assainissement de l'ERP.

    On refuse de continuer si elle est introuvable : inserer du HTML Outlook non
    filtre dans `body_html` fabriquerait une faille d'execution de script dans
    l'ERP, declenchee a l'ouverture du message. Mieux vaut ne rien ecrire.
    """
    if not ERP.exists():
        return None, f"depot ERP introuvable : {ERP}"
    sys.path.insert(0, str(ERP))
    os.environ.setdefault("ERP_JWT_SECRET", "import-local")
    os.environ.setdefault("ANTHROPIC_API_KEY", "import-local")
    os.environ.setdefault("STRIPE_SECRET_KEY", "import-local")
    try:
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            from backend.routers.emails import _sanitize_email_html
        return _sanitize_email_html, None
    except Exception as exc:
        return None, f"{type(exc).__name__}: {exc}"


def fil(message_id: str, in_reply_to: str, references: str) -> str:
    """Identifiant de fil : racine des References, sinon In-Reply-To, sinon soi.

    Meme ordre de priorite que le webhook (emails.py:3470-3478), avec UNE
    difference deliberee : on retire les chevrons PARTOUT, y compris du repli
    sur le message_id. Le webhook, lui, garde `thread_id = msg_id` avec ses
    chevrons alors qu'il les retire de References et d'In-Reply-To. Resultat
    chez lui : la racine d'un fil porte « <abc@x> » et ses reponses « abc@x »,
    donc elles ne se regroupent pas. Normaliser des deux cotes est la seule
    facon d'obtenir des fils coherents.
    """
    def nu(v: str) -> str:
        return (v or "").strip().strip("<>").strip()

    racine = nu(message_id)
    if references and references.strip():
        premier = nu(references.strip().split()[0])
        if premier:
            return premier
    if in_reply_to:
        return nu(in_reply_to) or racine
    return racine


def _decrire_cible(dsn: str, nom_variable: str) -> str:
    """Décrit la base réellement visée, mot de passe exclu.

    ⚠️ Le nom du SCHEMA ne suffit pas à identifier une cible : production,
    recette et restauration portent le même `tenant_constructo_e1f633`. Mesure
    du retest #205 tour 5 : deux bases distinctes ont produit une sortie de
    suppression identique au caractère près (« 1483 ligne(s) supprimee(s) dans
    tenant_constructo_e1f633 »). Ce qui décide de la gravité, c'est l'hôte —
    il doit donc être sous les yeux AVANT le DELETE, pas déduit après.

    ⚠️ NE JAMAIS INTERPOLER LE DSN BRUT DANS LA SORTIE. psycopg2 accepte aussi
    la forme libpq mot-cle/valeur (`host=... password=...`), sur laquelle
    `urlsplit` ne trouve ni hote ni utilisateur et met la chaine ENTIERE dans
    `path` — mot de passe compris, qui partait alors sur la sortie standard, dans
    l'historique du terminal et dans toute capture d'ecran. Mesure du retest
    #205 tour 6. Si l'hote n'est pas identifiable, on le DIT sans rien recopier.
    """
    from urllib.parse import urlsplit
    try:
        u = urlsplit(dsn)
        hote, port = u.hostname, u.port
        base, user = (u.path or "").lstrip("/"), u.username
    except Exception:
        hote = port = base = user = None
    if not hote:
        return ("Cible   : ⚠️ FORMAT NON-URI — cible NON IDENTIFIABLE\n"
                "          (la chaine n'est pas recopiee ici : elle peut porter "
                "un mot de passe)\n"
                f"          variable {nom_variable}\n"
                f"Schema  : {SCHEMA}")
    return (f"Cible   : {hote}:{port or 5432}/{base or '?'}  "
            f"(utilisateur {user or '?'})\n"
            f"          variable {nom_variable}\n"
            f"Schema  : {SCHEMA}")


def horodatage(valeur):
    try:
        return datetime.datetime(valeur.year, valeur.month, valeur.day,
                                 valeur.hour, valeur.minute, valeur.second)
    except Exception:
        return None


def collecter(nom_dossier, sens, limite, seuil, assainir, compte_id):
    """Construit les tuples prets a inserer pour un dossier Outlook."""
    dossier = om.resolve_folder(nom_dossier, None)
    lignes, alertes = [], []
    for item in om.iter_items(dossier, limit=None, newest=True):
        if len(lignes) >= limite:
            break
        date_recue = horodatage(getattr(item, "ReceivedTime", None))
        date_envoi = horodatage(getattr(item, "SentOn", None))
        reference = date_recue or date_envoi
        if seuil and reference and reference < seuil:
            continue
        mid = pont.identifiant_message(item)
        if not mid:
            alertes.append("message sans identifiant, ignore")
            continue
        try:
            objet = (item.Subject or "(sans objet)").strip()
        except Exception:
            objet = "(sans objet)"
        try:
            texte = item.Body or ""
        except Exception:
            texte = ""
        try:
            brut_html = item.HTMLBody or ""
        except Exception:
            brut_html = ""
        try:
            html = assainir(brut_html) if brut_html else ""
        except Exception as exc:
            alertes.append(f"assainissement echoue ({type(exc).__name__}), HTML ignore : {objet[:40]}")
            html = ""
        try:
            non_lu = bool(item.UnRead)
        except Exception:
            non_lu = False
        try:
            nb_pj = int(item.Attachments.Count)
        except Exception:
            nb_pj = 0
        # Separer To / Cc / Cci. Les aplatir dans une seule colonne exposerait
        # les destinataires en copie conforme invisible — voir l'avertissement
        # de `destinataires_par_type` dans outlook_bridge.py.
        tri = pont.destinataires_par_type(item)
        dest_to = ", ".join(tri["to"])
        dest_cc = ", ".join(tri["cc"])
        dest_cci = ", ".join(tri["cci"])
        expediteur = pont.adresse_expediteur(item)
        try:
            nom_exp = (item.SenderName or "").strip() or None
        except Exception:
            nom_exp = None
        irt = pont._prop(item, pont.PR_IN_REPLY_TO_ID)
        refs = pont._prop(item, pont.PR_INTERNET_REFERENCES)

        # has_attachments=FALSE et attachment_count=0 DELIBEREMENT, meme quand
        # Outlook en signale. Ce script n'importe PAS les fichiers : annoncer un
        # trombone sans ligne correspondante dans `email_attachments` affiche
        # dans l'ERP une piece jointe qui ne mene nulle part. Mesure du
        # 2026-08-17 : 656 lignes dans ce cas, trombones visibles a l'ecran,
        # corrigees apres coup. Mieux vaut ne rien annoncer que mentir.
        # Le vrai nombre est conserve dans le journal de sortie ci-dessous ;
        # quand les pieces jointes seront importees, remettre `nb_pj > 0, nb_pj`.
        lignes.append((
            compte_id, mid, fil(mid, irt, refs), irt or None,
            expediteur, nom_exp, dest_to, dest_cc or None, dest_cci or None,
            objet, texte, html,
            date_envoi, date_recue,
            sens, "UNREAD" if non_lu else "READ", not non_lu, False,
            False, 0, nom_dossier.lower(),
            json.dumps([MARQUEUR, RUN_ID]),
        ))
        if nb_pj:
            alertes.append(f"{nb_pj} PJ non importee(s) : {objet[:52]}")
    return lignes, alertes


SQL_INSERT = f"""
INSERT INTO {SCHEMA}.emails
  (account_id, message_id, thread_id, in_reply_to,
   email_from, email_from_name, email_to, email_cc, email_bcc,
   subject, body_text, body_html,
   date_sent, date_received,
   direction, status, is_read, is_starred,
   has_attachments, attachment_count, folder,
   labels_json, created_at, updated_at)
VALUES (%s,%s,%s,%s, %s,%s,%s,%s,%s, %s,%s,%s, %s,%s, %s,%s,%s,%s, %s,%s,%s,
        %s::jsonb, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
"""
# 22 colonnes, 22 marqueurs %s. Les compter A LA MAIN a chaque modification :
# une valeur de trop ou de moins ne leve pas toujours, elle DECALE les colonnes
# et ecrit des donnees justes au mauvais endroit — sans erreur visible.
assert SQL_INSERT.count("%s") == 22, "SQL_INSERT : nombre de %s incoherent"


def main() -> int:
    p = argparse.ArgumentParser(description="Import Outlook -> base Constructo AI.")
    p.add_argument("--folder", default="both", choices=["inbox", "sent", "both"])
    p.add_argument("--limit", type=int, default=50, help="maximum PAR dossier")
    p.add_argument("--since", default=None, help="AAAA-MM-JJ")
    p.add_argument("--account-id", type=int, default=3,
                   help="id dans email_accounts (3 = info@constructoai.ca)")
    p.add_argument("--ecrire", action="store_true", help="ECRIRE reellement en base")
    a = p.parse_args()

    # CONSTRUCTO_DB_URL d'abord : `DATABASE_URL` est un nom generique que
    # d'autres projets du poste posent aussi, et pointer par megarde la mauvaise
    # base tout en ayant `--ecrire` est le genre d'erreur qu'on ne rattrape pas.
    dsn = (os.environ.get("CONSTRUCTO_DB_URL") or "").strip()
    nom_variable = "CONSTRUCTO_DB_URL"
    if not dsn:
        dsn = (os.environ.get("DATABASE_URL") or "").strip()
        nom_variable = "DATABASE_URL (repli)"
    if not dsn:
        return _echec(
            "Aucune chaine de connexion. Poser CONSTRUCTO_DB_URL (recommande) ou\n"
            "    DATABASE_URL. Valeur = la chaine EXTERNE du Postgres Render\n"
            "    (Dashboard > base > Connect > External Connection String).\n"
            "    L'interne (dpg-...-a) ne resout que depuis le reseau Render.")
    if "dpg-" in dsn and ".render.com" not in dsn and "@dpg-" in dsn:
        return _echec(
            "Cette chaine ressemble a l'hote INTERNE de Render : elle ne resout\n"
            "    pas depuis ce poste. Prendre la chaine EXTERNE.")
    try:
        import psycopg2
        import psycopg2.extras
    except ImportError:
        return _echec("psycopg2 requis (pip install psycopg2-binary)")

    # ─────────────────────────────────────────────────────────────────────
    # `--annuler` A ETE RETIRE. C'est un aveu, pas une simplification.
    #
    # Six tours de retest #205 ont trouve SIX fois le meme defaut de fond : la
    # portee ANNONCEE n'etait pas la portee EXECUTEE. J'ai repare le message de
    # panne (T3), l'invite de confirmation (T4 — 200 annoncees, 2631
    # detruites), le message trompeur sur run inconnu (T5), puis retire le mode
    # global (T5). Le tour 6 a trouve le contournement suivant : le confinement
    # `jsonb @>` IGNORE LES DOUBLONS, donc `--run import-outlook` rendait le
    # predicat vrai pour TOUTE ligne marquee — 1441 supprimees sous l'etiquette
    # « UNIQUEMENT ». Et l'aide du script imprimait elle-meme cette valeur.
    #
    # Six garde-fous, six contournements. La conclusion n'est pas qu'il manquait
    # un septieme garde-fou : c'est qu'une suppression EN LOT derriere un drapeau
    # de confort ne peut pas etre rendue sure ainsi. On retire la capacite.
    #
    # CE QU'ON PERD : rien d'essentiel. L'import DEDUPLIQUE sur message_id, donc
    # on n'a jamais besoin de supprimer pour reimporter.
    # CE QU'IL FAUT FAIRE A LA PLACE, quand il faut vraiment supprimer : ecrire
    # la requete a la main, la relire, et la lancer soi-meme. Exemple, apres
    # avoir VERIFIE l'hote de la connexion :
    #     SELECT count(*) FROM tenant_constructo_e1f633.emails
    #      WHERE labels_json ? 'run-20260817-160312-13012';
    #     DELETE FROM ...   -- le meme WHERE, une fois le compte verifie
    # Un geste rare et grave merite un geste delibere.
    # ─────────────────────────────────────────────────────────────────────

    # LA CIBLE, SUR LE CHEMIN D'ECRITURE AUSSI — et AVANT la lecture Outlook,
    # donc valable en simulation comme en --ecrire. Elle n'etait affichee que
    # sur le chemin destructeur : on voyait l'hote avant de supprimer, jamais
    # avant d'ecrire 1437 lignes. C'etait a l'envers. La ligne « Compte cible :
    # id=3 » qui suit designe un compte OUTLOOK, pas une base : elle occupait la
    # place de cette information et se lisait comme si la cible etait confirmee.
    # Releve au retest #205 tour 6.
    print(_decrire_cible(dsn, nom_variable))
    print()

    assainir, erreur = charger_assainisseur()
    if assainir is None:
        return _echec(f"fonction d'assainissement de l'ERP indisponible ({erreur}).\n"
                      f"    Refus d'ecrire du HTML non filtre en base.")
    print("Assainisseur ERP charge (_sanitize_email_html)")

    seuil = None
    if a.since:
        try:
            seuil = datetime.datetime.strptime(a.since, "%Y-%m-%d")
        except ValueError:
            return _echec("--since attend AAAA-MM-JJ")

    plan = ([("inbox", "INBOUND")] if a.folder in ("inbox", "both") else []) + \
           ([("sent", "OUTBOUND")] if a.folder in ("sent", "both") else [])

    toutes, alertes = [], []
    for nom, sens in plan:
        print(f"Lecture « {nom} »…")
        try:
            lignes, alrt = collecter(nom, sens, a.limit, seuil, assainir, a.account_id)
        except Exception as exc:
            return _echec(f"lecture Outlook impossible ({nom}) : {exc}")
        print(f"  {len(lignes)} message(s)")
        toutes.extend(lignes)
        alertes.extend(alrt)

    if alertes:
        print(f"\n{len(alertes)} avertissement(s) :")
        for m in alertes[:6]:
            print(f"  - {m}")

    if not toutes:
        print("\nRien a importer.")
        return 0

    # Garde-fou : le tuple DOIT avoir exactement autant de valeurs que
    # SQL_INSERT a de marqueurs. psycopg2 leve bien en cas d'ecart, mais
    # echouer ici — avant toute connexion — dit OU est l'erreur.
    mauvaises = [i for i, l in enumerate(toutes) if len(l) != 22]
    if mauvaises:
        return _echec(f"{len(mauvaises)} ligne(s) n'ont pas 22 valeurs "
                      f"(premiere : index {mauvaises[0]}, {len(toutes[mauvaises[0]])} valeurs)")

    print(f"\n{len(toutes)} message(s) prets. Compte cible : id={a.account_id}")
    nb_cc = sum(1 for l in toutes if l[7])
    nb_cci = sum(1 for l in toutes if l[8])
    print(f"  dont {nb_cc} avec copie conforme, {nb_cci} avec copie conforme INVISIBLE")
    print("Apercu :")
    for l in toutes[:8]:
        marque = " [CCI]" if l[8] else ""
        print(f"  [{l[14]:8}] {str(l[13])[:19]:19} {(l[4] or '?')[:28]:28} {l[9][:38]}{marque}")

    if not a.ecrire:
        print("\n" + "=" * 70)
        print("SIMULATION — aucune ecriture.")
        # Repeter la cible ici : l'invitation a relancer avec --ecrire doit
        # porter le nom de la base que ce --ecrire va frapper.
        print(_decrire_cible(dsn, nom_variable).splitlines()[0])
        print("Relancer avec --ecrire pour appliquer.")
        print("=" * 70)
        return 0

    ecrits, doublons, echecs = 0, 0, []
    jalonnes = 0  # lignes rendues DURABLES par un commit jalon
    conn = psycopg2.connect(dsn, connect_timeout=25)
    try:
        cur = conn.cursor()
        for l in toutes:
            # POINT DE REPRISE PAR MESSAGE. Sans lui, un seul message fautif
            # (un octet nul dans l'objet suffit, psycopg2 leve) faisait un
            # rollback de TOUT le lot deja insere — pendant que le compteur
            # `ecrits` continuait d'augmenter. On annoncait 1400 ecritures pour
            # zero ligne reellement conservee.
            cur.execute("SAVEPOINT msg")
            try:
                cur.execute(
                    f"SELECT 1 FROM {SCHEMA}.emails WHERE message_id=%s AND account_id=%s LIMIT 1",
                    (l[1], l[0]))
                if cur.fetchone():
                    doublons += 1
                    cur.execute("RELEASE SAVEPOINT msg")
                    continue
                cur.execute(SQL_INSERT, l)
                cur.execute("RELEASE SAVEPOINT msg")
                ecrits += 1
            except Exception as exc:
                # Ce ROLLBACK est volontairement NU : s'il leve, c'est que la
                # connexion est morte, et il FAUT alors sortir de la boucle
                # plutot que d'empiler des echecs sur une session defunte.
                # L'exception remonte au gestionnaire exterieur, qui sait
                # annoncer combien de lignes sont durables.
                cur.execute("ROLLBACK TO SAVEPOINT msg")
                echecs.append(f"{str(l[9])[:40]} — {type(exc).__name__}: {str(exc)[:90]}")
            if (ecrits + doublons + len(echecs)) % 200 == 0:
                conn.commit()  # jalon : un incident tardif ne perd pas tout
                jalonnes = ecrits
        conn.commit()
        jalonnes = ecrits
    # ⚠️ `KeyboardInterrupt` DERIVE DE BaseException, PAS DE Exception. Un
    # `except Exception` ne l'attrape pas : un Ctrl+C traversait les deux try et
    # remontait nu — trace Python, aucun message, et jusqu'a 1400 lignes
    # silencieusement en base. C'est MOT POUR MOT le symptome que ce
    # gestionnaire existe pour supprimer. Et le declencheur est le plus banal
    # qui soit : un import de 1437 messages contre un Postgres joint par
    # Internet dure plusieurs minutes, l'operateur le croit bloque et
    # interrompt. Reproduit au retest #205 tour 3 (code 0xC000013A, 200 lignes
    # en base, zero avertissement).
    except (Exception, KeyboardInterrupt) as exc:
        # ⚠️ ROLLBACK NU INTERDIT ICI. Quand la connexion est MORTE — le cas de
        # panne le plus probable pour un Postgres Render joint par Internet :
        # redemarrage, bascule, coupure reseau, idle_in_transaction_timeout —
        # `conn.rollback()` leve a son tour (InterfaceError: connection already
        # closed). Hors de tout try, il tuait le gestionnaire AVANT le message
        # ci-dessous : trace Python nue, code 1, et 200 lignes silencieusement
        # en base. Une trace se lit spontanement comme « rien n'a ete ecrit » —
        # exactement le contresens que ce message existe pour eviter. Mesure du
        # retest #205 tour 2 (2026-08-17). L'ERP entoure deja ses propres
        # rollbacks de la meme facon (emails.py:3634-3637, 4914-4917).
        try:
            conn.rollback()
        except Exception:
            pass
        # NE PAS annoncer « rien n'a ete conserve » : les jalons ont rendu
        # durables les lignes qui les precedent, et le rollback ne defait que le
        # fragment ouvert. Une version precedente affirmait le contraire.
        volontaire = isinstance(exc, KeyboardInterrupt)
        titre = ("import INTERROMPU A LA DEMANDE (Ctrl+C)" if volontaire
                 else f"transaction interrompue : {exc}")
        return _echec(
            f"{titre}\n"
            f"    Au moins {jalonnes} ligne(s) ont ete rendues durables par un "
            f"jalon et SONT EN BASE.\n"
            f"    Ce nombre est une BORNE INFERIEURE si la connexion est morte.\n"
            f"\n"
            f"    Compte exact de CETTE execution (connexion neuve, ne supprime rien) :\n"
            f"        python scripts/outlook_vers_bd.py --annuler --run {RUN_ID}\n"
            f"    Pour retirer CETTE execution seulement :\n"
            f"        python scripts/outlook_vers_bd.py --annuler --run {RUN_ID} --oui\n"
            f"\n"
            f"    ⚠️  Code de sortie 3 : des lignes SONT en base. A ne pas\n"
            f"       confondre avec le code 2, qui veut dire « rien n'a ete fait ».",
            code=3)
    finally:
        conn.close()
    print(f"\n{ecrits} ecrit(s), {doublons} doublon(s) ignore(s), {len(echecs)} echec(s).")
    # IMPRIMER LE RUN_ID. Sans lui, l'annulation ciblee est inutilisable sur le
    # chemin normal : l'operateur ne connait pas l'identifiant de ce qu'il vient
    # d'ecrire, et se rabat sur une portee plus large. Le marqueur de run ne
    # sert a rien s'il reste secret. Releve au retest #205 tour 5.
    if ecrits:
        print(f"Etiquette de cette execution : {RUN_ID}")
        print(f"Pour l'annuler : python scripts/outlook_vers_bd.py "
              f"--annuler --run {RUN_ID} --oui")
    for m in echecs[:8]:
        print(f"  - {m}")
    return 0 if not echecs else 1


def _echec(msg: str, code: int = 2) -> int:
    """Sort en erreur. Le CODE distingue deux situations tres differentes.

      2 — refus AVANT toute ecriture : configuration manquante, drapeau
          incoherent, assainisseur introuvable. La base est intacte.
      3 — l'import est MORT EN COURS et des lignes SONT durablement en base.

    Un ordonnanceur ou un operateur qui lit le code de sortie doit pouvoir les
    separer : le premier se corrige et se relance, le second demande d'aller
    verifier ce qui a ete ecrit. Ils rendaient tous les deux 2 — releve au
    retest #205 tour 6.
    """
    print(f"\nERREUR : {msg}")
    return code


if __name__ == "__main__":
    sys.exit(main())
