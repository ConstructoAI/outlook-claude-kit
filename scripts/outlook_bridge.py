#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Pont local Outlook -> Constructo AI.

CE QUE C'EST. Un prototype qui lit la boite Outlook DEJA AUTHENTIFIEE du poste
(MAPI/COM, comme `outlook_mail.py`) et pousse les messages vers l'endpoint
`POST /api/erp/v1/emails/webhook/inbound` de l'ERP.

POURQUOI. Microsoft a retire l'authentification de base d'Exchange Online (IMAP
fin 2022, SMTP AUTH le 30 avril 2026) : le module courriel de l'ERP ne peut plus
se connecter par mot de passe. La voie OAuth exige une inscription d'application
Entra ID, or les locataires Microsoft 365 revendus par GoDaddy ne donnent pas les
droits d'administrateur necessaires. Ce pont contourne le probleme entierement :
il ne demande aucun mot de passe, aucun jeton, aucun consentement — il se branche
sur la session Outlook qui fonctionne deja.

SECURITE PAR DEFAUT. Sans `--push`, le script N'ENVOIE RIEN. Il affiche ce qu'il
pousserait et ecrit les charges utiles dans un fichier pour inspection. Le secret
du webhook n'est jamais passe en argument de ligne de commande (il serait visible
dans l'historique du shell et dans la liste des processus) : il est lu dans la
variable d'environnement CONSTRUCTO_WEBHOOK_SECRET.

ROUTAGE — LE PIEGE. `_resolve_tenant_from_email` cote serveur resout le locataire
a partir de la PARTIE LOCALE de l'adresse destinataire, comparee au `slug` de
`public.entreprises`. Un message adresse a `info@constructoai.ca` ne route donc
vers RIEN (sauf si un locataire porte le slug « info »). Il faut pousser vers
`{slug}@constructoai.ca`, d'ou l'option --tenant-slug, obligatoire.
Les vrais destinataires du message sont conserves dans `internetMessageHeaders`,
que le serveur relit en priorite — donc l'affichage reste juste.

IDEMPOTENCE. Le serveur deduplique sur `message_id` (index unique). Ce script
tient EN PLUS un journal local pour eviter de renvoyer 800 messages a chaque
execution. Supprimer le fichier d'etat force un renvoi complet, sans risque de
doublon cote serveur.

Usage :
    python scripts/outlook_bridge.py --tenant-slug constructoai --limit 20
    python scripts/outlook_bridge.py --tenant-slug constructoai --since 2026-08-01 \
        --url https://app.constructoai.ca --push
"""

from __future__ import annotations

import argparse
import base64
import datetime
import hashlib
import json
import os
import sys
import tempfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ICI = Path(__file__).resolve().parent
sys.path.insert(0, str(ICI))

try:
    import outlook_mail as om
except ImportError as exc:  # pragma: no cover - depend du poste
    sys.exit(f"outlook_mail.py introuvable a cote de ce script : {exc}")

# Proprietes MAPI nommees. Elles n'ont pas d'equivalent dans l'objet COM de haut
# niveau et doivent passer par PropertyAccessor.
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
PR_IN_REPLY_TO_ID = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"
PR_INTERNET_REFERENCES = "http://schemas.microsoft.com/mapi/proptag/0x1039001F"

ETAT_DEFAUT = ICI.parent / ".bridge_state.json"
TAILLE_PJ_MAX = 8 * 1024 * 1024  # 8 Mo par piece jointe


# ------------------------------------------------------------------ utilitaires
def _prop(item, schema_uri, defaut=""):
    """Lit une propriete MAPI nommee. Retourne `defaut` si absente."""
    try:
        valeur = item.PropertyAccessor.GetProperty(schema_uri)
        return (valeur or "").strip() if isinstance(valeur, str) else (valeur or defaut)
    except Exception:
        return defaut


def adresse_expediteur(item) -> str:
    """Adresse SMTP de l'expediteur.

    SenderEmailAddress rend une adresse Exchange (/O=EXCHANGELABS/OU=...) pour
    tout envoi interne — 26 messages de la boite mesuree etaient dans ce cas.
    Telle quelle, elle est inexploitable cote ERP (le rattachement au CRM se fait
    par adresse). On passe donc par PR_SENDER_SMTP_ADDRESS, puis par l'annuaire
    Exchange, avant de se rabattre sur la valeur brute.
    """
    smtp = _prop(item, PR_SENDER_SMTP_ADDRESS)
    if smtp and "@" in smtp:
        return smtp
    try:
        entree = item.Sender
        if entree is not None:
            try:
                utilisateur = entree.GetExchangeUser()
                if utilisateur and utilisateur.PrimarySmtpAddress:
                    return utilisateur.PrimarySmtpAddress
            except Exception:
                pass
            depuis_prop = _prop(entree, PR_SMTP_ADDRESS)
            if depuis_prop and "@" in depuis_prop:
                return depuis_prop
    except Exception:
        pass
    try:
        brut = (item.SenderEmailAddress or "").strip()
    except Exception:
        brut = ""
    return brut


def _adresse_destinataire(r) -> str:
    """Adresse SMTP d'un destinataire, avec replis successifs."""
    adr = _prop(r, PR_SMTP_ADDRESS)
    if adr and "@" in adr:
        return adr
    for attribut in ("Address", "Name"):
        try:
            v = (getattr(r, attribut) or "").strip()
        except Exception:
            v = ""
        if v:
            return v
    return ""


def destinataires_par_type(item) -> dict:
    """Destinataires SMTP separes en To / Cc / Cci.

    ⚠️ LIRE `r.Type` EST OBLIGATOIRE. `item.Recipients` melange les trois types
    (olTo=1, olCC=2, olBCC=3) dans une seule collection, et les destinataires en
    copie conforme INVISIBLE y figurent tels quels sur les elements du dossier
    Envoyes. Une version anterieure de ce code ne lisait pas Type et concatenait
    tout dans `email_to` : mesure du 2026-08-17 sur la vraie boite — 74 messages
    du dossier Envoyes portaient au moins un destinataire Cci, dont un envoi de
    prospection a **754 adresses**, ecrit en base sur 19 318 caracteres comme
    une liste VISIBLE. Des adresses volontairement masquees les unes aux autres
    devenaient consultables. Les 1 426 lignes concernees ont ete supprimees.
    """
    tri = {"to": [], "cc": [], "cci": []}
    vus = 0
    try:
        for r in item.Recipients:
            vus += 1
            adr = _adresse_destinataire(r)
            if not adr:
                continue
            try:
                t = int(r.Type)
            except Exception:
                t = 1  # type illisible : le traiter en To, jamais en Cci
            if t == 3:
                tri["cci"].append(adr)
            elif t == 2:
                tri["cc"].append(adr)
            else:
                tri["to"].append(adr)
    except Exception:
        pass
    # Repli quand la collection est vide ou illisible. Il n'alimente QUE `to` :
    # les proprietes item.CC / item.BCC sont des chaines d'affichage peu fiables,
    # et se tromper dans ce sens-la exposerait des adresses masquees.
    if vus == 0:
        try:
            brut = (item.To or "").strip()
            if brut:
                tri["to"] = [p.strip() for p in brut.split(";") if p.strip()]
        except Exception:
            pass
    return tri


def adresses_destinataires(item) -> list[str]:
    """Destinataires VISIBLES uniquement (To + Cc). N'expose jamais les Cci."""
    tri = destinataires_par_type(item)
    return tri["to"] + tri["cc"]


def identifiant_message(item) -> str:
    """Message-Id Internet, ou un substitut STABLE derive de l'EntryID.

    L'identifiant sert de cle de deduplication cote serveur. Un identifiant qui
    changerait d'une execution a l'autre creerait un doublon a chaque passage :
    c'est pour cela qu'on hache l'EntryID plutot que d'inventer un UUID.
    """
    mid = _prop(item, PR_INTERNET_MESSAGE_ID)
    if mid:
        return mid
    try:
        graine = item.EntryID or ""
    except Exception:
        graine = ""
    if not graine:
        return ""
    empreinte = hashlib.sha256(graine.encode("utf-8", "replace")).hexdigest()[:32]
    return f"<{empreinte}@pont-outlook.constructoai.ca>"


def pieces_jointes(item, activer: bool) -> tuple[list[dict], list[str]]:
    """Extrait les pieces jointes en base64. Retourne (liste, avertissements)."""
    if not activer:
        return [], []
    sorties, alertes = [], []
    dossier = Path(tempfile.mkdtemp(prefix="pont_outlook_"))
    try:
        for piece in item.Attachments:
            try:
                nom = (piece.FileName or "piece").strip()
            except Exception:
                nom = "piece"
            chemin = dossier / nom.replace("/", "_").replace("\\", "_")
            try:
                piece.SaveAsFile(str(chemin))
                donnees = chemin.read_bytes()
            except Exception as exc:
                alertes.append(f"piece jointe illisible ({nom}) : {type(exc).__name__}")
                continue
            finally:
                try:
                    chemin.unlink(missing_ok=True)
                except Exception:
                    pass
            if len(donnees) > TAILLE_PJ_MAX:
                alertes.append(
                    f"piece jointe ignoree, {len(donnees) // 1024} Ko > "
                    f"{TAILLE_PJ_MAX // 1024} Ko : {nom}"
                )
                continue
            sorties.append({
                "filename": nom[:200],
                "contentType": "application/octet-stream",
                "data": base64.b64encode(donnees).decode("ascii"),
            })
    except Exception as exc:
        alertes.append(f"enumeration des pieces jointes interrompue : {type(exc).__name__}")
    finally:
        try:
            dossier.rmdir()
        except Exception:
            pass
    return sorties, alertes


def charge_utile(item, slug: str, avec_pj: bool) -> tuple[dict, list[str]]:
    """Construit la charge utile JSON attendue par /webhook/inbound."""
    reels = adresses_destinataires(item)
    try:
        objet = (item.Subject or "(sans objet)").strip()
    except Exception:
        objet = "(sans objet)"
    try:
        texte = item.Body or ""
    except Exception:
        texte = ""
    try:
        html = item.HTMLBody or ""
    except Exception:
        html = ""

    pj, alertes = pieces_jointes(item, avec_pj)

    charge = {
        # Adresse de ROUTAGE : la partie locale doit egaler le slug du locataire.
        "to": [f"{slug}@constructoai.ca"],
        # ⚠️ NE PAS ENVOYER `internetMessageHeaders`. Le serveur ne se contente
        # pas de le LIRE : sa logique de « preservation d'alias M365 »
        # (emails.py:3316-3330) ECRASE le destinataire de routage par ce qu'elle
        # y trouve. Mesure du 2026-08-17 : en y mettant les vrais destinataires,
        # les 3 messages ont ete acceptes (HTTP 200) puis JETES —
        # « no tenant for recipient ci_activity@noreply.github.com ». Le pont
        # annoncait 3/3 reussis et la base n'avait rien recu.
        # Consequence assumee : `email_to` portera l'adresse de routage et non
        # les destinataires d'origine. Pour du courrier ENTRANT vers le
        # locataire, c'est acceptable ; le corps du message garde le detail.
        "from": adresse_expediteur(item),
        "subject": objet,
        "text": texte,
        "html": html,
        "messageId": identifiant_message(item),
        "inReplyTo": _prop(item, PR_IN_REPLY_TO_ID),
        "references": _prop(item, PR_INTERNET_REFERENCES),
        "attachments": pj,
    }
    return charge, alertes


# ------------------------------------------------------------------------ etat
def lire_etat(chemin: Path) -> set[str]:
    try:
        return set(json.loads(chemin.read_text(encoding="utf-8")).get("pousses", []))
    except Exception:
        return set()


def ecrire_etat(chemin: Path, deja: set[str]) -> None:
    try:
        chemin.write_text(
            json.dumps({"pousses": sorted(deja)}, ensure_ascii=False, indent=1),
            encoding="utf-8",
        )
    except Exception as exc:
        print(f"  [!] etat non enregistre ({type(exc).__name__}) — "
              f"la prochaine execution reproposera ces messages")


# ------------------------------------------------------------------------ main
def main() -> int:
    p = argparse.ArgumentParser(
        description="Pont local Outlook -> Constructo AI (simulation par defaut).")
    p.add_argument("--tenant-slug", required=True,
                   help="slug du locataire dans public.entreprises (adresse de routage)")
    p.add_argument("--folder", default="inbox", help="dossier Outlook (defaut : inbox)")
    p.add_argument("--account", default=None, help="adresse du compte si le profil en a plusieurs")
    p.add_argument("--limit", type=int, default=25, help="nombre maximum de messages")
    p.add_argument("--since", default=None, help="ne prendre que les messages recus depuis AAAA-MM-JJ")
    p.add_argument("--attachments", action="store_true", help="inclure les pieces jointes en base64")
    p.add_argument("--url", default=os.environ.get("CONSTRUCTO_URL", "https://app.constructoai.ca"),
                   help="base de l'ERP")
    p.add_argument("--push", action="store_true",
                   help="ENVOYER reellement. Sans ce drapeau, rien ne quitte le poste.")
    p.add_argument("--state", default=str(ETAT_DEFAUT), help="fichier de journal local")
    p.add_argument("--dump", default=None, help="ecrire les charges utiles dans ce fichier JSON")
    p.add_argument("--rejouer", action="store_true",
                   help="ignorer le journal local et reproposer les messages deja pousses")
    a = p.parse_args()

    secret = os.environ.get("CONSTRUCTO_WEBHOOK_SECRET", "").strip()
    if a.push and not secret:
        return _echec("--push demande mais CONSTRUCTO_WEBHOOK_SECRET est absente de "
                      "l'environnement.\n    Sous PowerShell :  $env:CONSTRUCTO_WEBHOOK_SECRET "
                      "= '<valeur de N8N_WEBHOOK_SECRET>'")

    chemin_etat = Path(a.state)
    deja = set() if a.rejouer else lire_etat(chemin_etat)

    print(f"Lecture du dossier « {a.folder} »…")
    try:
        dossier = om.resolve_folder(a.folder, a.account)
        items = om.iter_items(dossier, limit=None, unread=False, newest=True)
    except Exception as exc:
        return _echec(f"lecture Outlook impossible : {exc}\n"
                      f"    Outlook est-il lance ? (python scripts/check_setup.py)")

    seuil = None
    if a.since:
        try:
            seuil = datetime.datetime.strptime(a.since, "%Y-%m-%d")
        except ValueError:
            return _echec(f"--since attend le format AAAA-MM-JJ, recu : {a.since}")

    charges, ignores_date, ignores_deja, alertes = [], 0, 0, []
    for item in items:
        if len(charges) >= a.limit:
            break
        if seuil is not None:
            try:
                recu = item.ReceivedTime
                if datetime.datetime(recu.year, recu.month, recu.day) < seuil:
                    ignores_date += 1
                    continue
            except Exception:
                pass
        mid = identifiant_message(item)
        if not mid:
            alertes.append("message sans identifiant exploitable, ignore")
            continue
        if mid in deja:
            ignores_deja += 1
            continue
        charge, sous_alertes = charge_utile(item, a.tenant_slug, a.attachments)
        alertes.extend(sous_alertes)
        charges.append(charge)

    print(f"  {len(charges)} message(s) retenu(s)"
          f"{f', {ignores_deja} deja pousse(s)' if ignores_deja else ''}"
          f"{f', {ignores_date} hors periode' if ignores_date else ''}")
    if alertes:
        print(f"  {len(alertes)} avertissement(s) :")
        for m in alertes[:8]:
            print(f"    - {m}")

    if not charges:
        print("\nRien a faire.")
        return 0

    print(f"\nRoutage    : {a.tenant_slug}@constructoai.ca")
    print(f"Destination: {a.url.rstrip('/')}/api/erp/v1/emails/webhook/inbound")
    print("\nApercu :")
    for c in charges[:10]:
        npj = len(c["attachments"])
        print(f"  · {(c['from'] or '(inconnu)')[:34]:34} "
              f"{c['subject'][:56]}{f'  [{npj} PJ]' if npj else ''}")
    if len(charges) > 10:
        print(f"  … et {len(charges) - 10} autre(s)")

    if a.dump:
        Path(a.dump).write_text(json.dumps(charges, ensure_ascii=False, indent=2),
                                encoding="utf-8")
        print(f"\nCharges utiles ecrites dans {a.dump}")

    if not a.push:
        print("\n" + "=" * 70)
        print("SIMULATION — rien n'a quitte ce poste.")
        print("Pour envoyer reellement, relancer avec --push")
        print("=" * 70)
        return 0

    try:
        import requests
    except ImportError:
        return _echec("le module `requests` est requis pour --push (pip install requests)")

    cible = f"{a.url.rstrip('/')}/api/erp/v1/emails/webhook/inbound"
    entetes = {"Authorization": f"Bearer {secret}", "Content-Type": "application/json"}
    ok, jetes, echecs = 0, 0, []
    print(f"\nEnvoi vers {cible} …")
    for i, c in enumerate(charges, 1):
        try:
            r = requests.post(cible, json=c, headers=entetes, timeout=45)
            if 200 <= r.status_code < 300:
                # ⚠️ UN 200 NE PROUVE RIEN. Quand aucun locataire ne correspond
                # au destinataire, le serveur journalise « no tenant for
                # recipient » et repond QUAND MEME un succes, sans rien stocker.
                # Mesure du 2026-08-17 : 3/3 « acceptes », 0 ligne en base. On
                # lit donc le compte reel dans la reponse.
                try:
                    corps = r.json()
                except Exception:
                    corps = {}
                livres = corps.get("delivered", corps.get("delivered_count"))
                # FERMETURE PAR DEFAUT. Ne compter un succes QUE sur une preuve
                # explicite `delivered >= 1`. Un champ absent doit compter comme
                # un echec : le serveur a DEUX chemins qui rendent 200 avec un
                # corps vide — le destinataire introuvable (emails.py:3379) et
                # surtout le rattrapage general du webhook (emails.py:3651,
                # « Retourner 200 quand meme — sinon Mailgun retry »), qui se
                # declenche sur n'importe quelle panne de base. Une version
                # precedente de ce correctif traitait l'absence comme un succes
                # ET gravait le messageId dans le journal local : 800 messages
                # pouvaient etre annonces stockes, perdus, et jamais reproposes.
                # Trouve par le retest #205 du 2026-08-17.
                try:
                    nb = int(livres)
                except (TypeError, ValueError):
                    nb = -1
                if nb < 1:
                    jetes += 1
                    motif = ("aucun locataire pour l'adresse de routage"
                             if nb == 0 else
                             "reponse sans champ `delivered` — panne serveur probable")
                    echecs.append(f"{c['subject'][:40]} — NON STOCKE ({motif})")
                    continue
                ok += 1
                deja.add(c["messageId"])
            else:
                echecs.append(f"{c['subject'][:40]} — HTTP {r.status_code} {r.text[:110]}")
        except Exception as exc:
            echecs.append(f"{c['subject'][:40]} — {type(exc).__name__}: {exc}")
        if i % 10 == 0:
            print(f"  {i}/{len(charges)}…")
            ecrire_etat(chemin_etat, deja)  # sauvegarde intermediaire

    ecrire_etat(chemin_etat, deja)
    print(f"\n{ok}/{len(charges)} message(s) reellement stocke(s) par l'ERP.")
    if jetes:
        print(f"⚠️  {jetes} accepte(s) mais NON stocke(s) — verifier --tenant-slug "
              f"contre public.entreprises.slug")
    if echecs:
        print(f"{len(echecs)} probleme(s) :")
        for m in echecs[:10]:
            print(f"  - {m}")
        return 1
    return 0


def _echec(message: str) -> int:
    print(f"\nERREUR : {message}")
    return 2


if __name__ == "__main__":
    sys.exit(main())
