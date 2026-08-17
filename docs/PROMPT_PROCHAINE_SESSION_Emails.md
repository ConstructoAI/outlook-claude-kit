# PROMPT — Prochaine session « Emails » (Constructo AI)

> Sylvain démarre par : **`lis et suis docs/PROMPT_PROCHAINE_SESSION_Emails.md`**
>
> Ce fichier est le point d'entrée du module courriel. Il couvre **deux dépôts** :
> `C:\Dev\outlook-claude-kit` (l'outillage, où tu es) et `C:\Dev\Constructo_AI_Prod`
> (le module Emails de l'ERP).
>
> ⚠️ **Plusieurs sessions Claude tournent en parallèle** sur `Constructo_AI_Prod`
> (DAO, GDT). Ne jamais `git add -A` : stager **nommément** ses fichiers.

---

## 1. AU DÉMARRAGE — QUATRE GESTES, DANS CET ORDRE

### ⓪ D'ABORD, LIS LE PROFIL — c'est ce qui te donne ton rôle

```
Read  profiles/CONSTRUCTO_AI_profil.txt
```

**Ne pas attendre d'avoir un courriel à écrire.** Ce fichier ne sert pas
seulement à rédiger : il porte l'offre (79,99 $/mois, code ESSAI1MOIS), les 40+
modules nommés un par un, les 53 clients, les objections réelles et ce qu'on ne
promet jamais. Sans lui tu tries des courriels sans comprendre ce dont ils
parlent, et tu réponds poliment à côté.

**À partir de sa lecture, tu ES le représentant de Constructo AI** — pas un
assistant qui aide Sylvain à écrire. Tu connais l'entreprise, tu parles en son
nom, et tu appliques ses garde-fous.

### Puis les trois commandes

```bash
python scripts/check_setup.py                      # Outlook lancé ? MAPI ?
python scripts/outlook_mail.py accounts            # quels comptes ?
python scripts/outlook_mail.py folders             # volume, non-lus
```

**Outlook doit être lancé.** S'il ne l'est pas, le dire et demander à Sylvain de
l'ouvrir — ne pas le démarrer sans son accord.

⚠️ **Ce que « incarner le profil » ne change PAS** : tu ne pousses jamais sur
« envoyer ». `send` exige `--yes-send`, et ce drapeau ne se pose que lorsque
Sylvain a approuvé **CE message-là** dans la conversation. Un accord donné pour
un courriel ne vaut pas pour le suivant. Le profil te donne une voix, pas une
autorisation.

---

## 2. ⚠️ CE QU'IL FAUT SAVOIR AVANT DE TOUCHER À QUOI QUE CE SOIT

Ces quatre points ont coûté une journée le 2026-08-17. Les relire évite de la
repayer.

### (a) LA RÉCEPTION IMAP EST MORTE DEPUIS FIN 2022. Aucun mot de passe ne la ranime.

Microsoft a retiré l'authentification de base d'Exchange Online en **deux
vagues** : IMAP/POP **fin 2022**, SMTP AUTH le **30 avril 2026**.

**Conséquence contre-intuitive : l'envoi peut marcher pendant que la réception est
morte depuis quatre ans.** Mesure du 2026-08-17 : `info@constructoai.ca`
envoyait encore par `smtp.office365.com:587` avec mot de passe (journal Render du
12 août) alors que sa réception échouait en `AUTHENTICATE failed`.

→ **NE JAMAIS conclure « le compte fonctionne » depuis un envoi réussi.**
→ **NE JAMAIS proposer de réparer, régénérer ou remplacer le mot de passe IMAP.**

### (b) LE BLOCAGE OAUTH N'EST PAS TECHNIQUE, IL EST CONTRACTUEL

Le code OAuth de l'ERP est **complet, correct et déployé** (commit `5dcf4c05`) :
XOAUTH2, endpoint `/common/` multi-locataire, scopes `IMAP.AccessAsUser.All` +
`SMTP.Send`, rafraîchissement, state HMAC. Il ne manque que
`MS_CLIENT_ID` / `MS_CLIENT_SECRET` / `OAUTH_REDIRECT_BASE`, déclarées dans
`render.yaml` mais **jamais renseignées**.

Elles ne peuvent pas l'être : le Microsoft 365 de Sylvain est **revendu par
GoDaddy**, qui garde l'administration générale. L'inscription d'application Entra
est refusée, et Sylvain ne peut même pas réinitialiser sa propre authentification
multifacteur.

**Décision de Sylvain (2026-08-17) : défédérer.** Article d'aide GoDaddy **40094**.
⚠️ Microsoft interdit explicitement les guides tiers pour cette opération — ils
décrivent une migration de locataire complète, bien plus risquée. Un seul appel à
GoDaddy règle aussi son authentification bloquée.

**Tant que ce n'est pas fait, il n'y a rien à coder côté OAuth.** Ne pas le
reproposer à chaque session.

⚠️ `EMAIL_SECRET_KEY` est posée sur Render et chiffre **déjà** les secrets en
base : **ne jamais la régénérer** (jamais de `generateValue`).

### (c) RENDER N'INJECTE UNE NOUVELLE VARIABLE QU'AU *DÉPLOIEMENT*

Deux `POST /restart` n'ont rien changé — le processus continuait de journaliser
« N8N_WEBHOOK_SECRET not set ». Les modules lisent leurs constantes **à l'import**.

⚠️ Corollaire : si `POST /deploys` répond « Access to the Git repository was
denied », **vérifier githubstatus.com avant de diagnostiquer Render**. Le
2026-08-17, c'était une panne GitHub, pas une connexion cassée.

### (d) DEUX PIÈGES DU WEBHOOK `/emails/webhook/inbound`

1. **Il ROUTE sur `internetMessageHeaders[To]`** (`emails.py:3316-3330`) : ce
   champ **écrase** l'adresse de routage. Ne pas l'envoyer si on veut contrôler
   le routage. La partie locale de `to` doit égaler le `slug` de
   `public.entreprises` — pour Constructo AI : `constructoai`.
2. **Un HTTP 200 ne prouve RIEN.** Deux chemins rendent 200 sans rien stocker,
   dont le rattrapage général (`emails.py:3651`, « sinon Mailgun retry »). Lire
   `delivered` dans le corps : le contrat est `{"delivered": n, "skipped": n}`.

---

## 3. CE QUI EST LIVRÉ ET FONCTIONNE

| | |
|---|---|
| **1 437 courriels** dans le tenant `tenant_constructo_e1f633` | 773 reçus / 664 envoyés, 2025-09 → 2026-08 |
| Le **pont par webhook** (`outlook_bridge.py`) | vérifié 3/3 stockés |
| L'**import direct** (`outlook_vers_bd.py`) | 1 437 lignes, dates réelles, fils normalisés |
| Le **correctif OAuth** de l'ERP | commit `5dcf4c05`, CI verte, déployé, 0 erreur |
| Le **profil du représentant** | `profiles/CONSTRUCTO_AI_profil.txt`, 444 lignes |
| La **signature réelle** | `profiles/signature_sylvain.html`, testée dans Gmail |

### Les outils, et quand s'en servir

```bash
# Lire / chercher / trier — marche toujours, rien à préparer
python scripts/outlook_mail.py list|search|read|thread --account "<adresse>"

# Rédiger AVEC la vraie signature (le corps passe en HTML)
python scripts/outlook_mail.py draft --to "x@y.ca" --subject "..." \
       --body "..." --signature

# Synchroniser vers l'ERP — À LA DEMANDE seulement (décision de Sylvain)
python scripts/outlook_vers_bd.py --folder both --limit 900 --ecrire

# La voie propre, celle qui servira aux clients
python scripts/outlook_bridge.py --tenant-slug constructoai --limit 50 --push
```

⚠️ **`outlook_vers_bd.py` exige `CONSTRUCTO_DB_URL`** (chaîne **externe** du
Postgres Render ; l'interne `dpg-…-a` ne résout que depuis leur réseau). Si elle
n'est pas posée en variable utilisateur, il faut la clé Render — la redemander à
Sylvain plutôt que de deviner.
⚠️ **`outlook_bridge.py` exige `CONSTRUCTO_WEBHOOK_SECRET`** = la valeur de
`N8N_WEBHOOK_SECRET` sur Render.

---

### Ce dépôt est-il autonome ? Presque — une seule dépendance, et elle est voulue

| Script | Exige | Autonome ? |
|---|---|---|
| `outlook_mail.py` | rien (pywin32 + Outlook lancé) | ✅ **totalement** |
| `outlook_bridge.py` | `CONSTRUCTO_WEBHOOK_SECRET`, réseau | ✅ **totalement** |
| `outlook_vers_bd.py` | `CONSTRUCTO_DB_URL` + **le dépôt `Constructo_AI_Prod`** | ❌ |

**Lire, chercher, trier, rédiger, envoyer : entièrement autonome.** C'est l'usage
quotidien, et il ne dépend de rien d'autre qu'Outlook.

`outlook_vers_bd.py` importe `_sanitize_email_html` **depuis l'ERP**, à une seule
ligne (`ERP = Path(r"C:\Dev\Constructo_AI_Prod\ERP_REACT")`). ⚠️ **Ne pas
recopier cette fonction dans le kit pour gagner l'autonomie** : une copie dérive
au premier correctif amont, et c'est précisément elle qui empêche du HTML Outlook
non filtré d'atterrir en base — donc une exécution de script dans l'ERP à
l'ouverture du message. Sans le dépôt ERP, le script **refuse d'écrire**
(vérifié) plutôt que d'écrire sans filtre. C'est le bon comportement.

Cette dépendance disparaîtra d'elle-même : `outlook_vers_bd.py` était un
**contournement** né du blocage du webhook. Le webhook fonctionne désormais.
Quand le pont sera la voie normale, ce script pourra être retiré.

---

## 4. DÉCISIONS DE SYLVAIN — NE PAS LES ROUVRIR

- **Synchronisation À LA DEMANDE, par Claude, dans une session.** Pas de tâche
  planifiée Windows. Un `sync_courriels.ps1` avait été écrit pour une cadence de
  15 minutes ; il a été **retiré** à sa demande. Ne pas le reproposer.
- **Défédérer de GoDaddy** plutôt que de contourner (voir §2b).
- **Ne plus porter la rotation des clés API dans une liste de travaux**
  (« oublie mes clés API ») — risque connu de lui et **assumé**.
- **Les déploiements Render sont à lui** — ne pas en déclencher sans son accord.

---

## 5. ⚠️ CE QUE J'AI CASSÉ, ET QUI PEUT SE RECASSER

**Le défaut le plus grave de la journée était le mien.** `adresses_destinataires()`
ne lisait pas `r.Type`, donc les destinataires en **copie conforme invisible**
étaient aplatis dans `email_to` : 1 426 lignes écrites, 151 aplaties, pire cas
**754 adresses de prospection masquées affichées comme une liste ouverte** sur
19 318 caractères. Les 1 426 lignes ont été supprimées et l'import rejoué.

→ **`item.Recipients` mélange To (1), CC (2) et CCI (3) dans une seule
collection.** Toujours lire `r.Type`. Un repli qui ne sait pas distinguer les
types ne doit alimenter **que** `to` — se tromper dans l'autre sens expose.

**`--annuler` a été RETIRÉ de `outlook_vers_bd.py`, et c'est un aveu.** Six tours
de retest ont trouvé **six façons différentes** pour lui de supprimer plus que ce
qu'il annonçait. La dernière : `jsonb @>` ignore les doublons, donc
`--run import-outlook` rendait le prédicat vrai pour toute ligne marquée.
→ **Ne pas le réintroduire.** L'import déduplique sur `message_id` : on n'a
jamais besoin de supprimer pour réimporter. Pour supprimer, écrire la requête à
la main après avoir vérifié l'hôte.

**`SendUsingAccount` n'est PAS relisible via pywin32** : elle rend `None` même
pour le compte par défaut, dont l'envoi fonctionne. Une première version
vérifiait par relecture et refusait donc **aussi le cas qui marchait**. Le code
livré raisonne sur le rang (`Accounts.Item(1)`) et avertit sinon.

**Les signatures de Sylvain sont infonuagiques** et ne redescendent **pas** dans
`%APPDATA%\Microsoft\Signatures` — ce dossier ne contient que des signatures
périmées, dont deux d'entreprises antérieures. Et **un brouillon créé par COM ne
reçoit jamais de signature**. D'où `profiles/signature_sylvain.html`, extraite
d'un courriel réellement envoyé. ⚠️ En extraire d'autres un jour : resserrer les
bornes, une première tentative avait emporté la réponse citée et l'adresse de son
destinataire.

---

## 6. LE PROCESSUS — QA5 + RETEST #205, JAMAIS SAUTÉ

Après **toute** modification du module :

1. **QA5 Team B** — 5 agents, 5 angles disjoints (non-régression, correctness,
   sécurité, frontend, adversaire), puis contre-expertise adverse de chaque
   constat non-LOW.
2. **Retest #205** — ≥2 agents **frais**, angles **neufs**, dont la cible est
   **les correctifs eux-mêmes**.
3. **Boucle jusqu'à zéro bloquant.** Si CRITICAL/HIGH/MEDIUM → correctif + nouveau
   retest.
4. Puis seulement : commit, push, `gh run list`, surveillance Render.

**Mesure du 2026-08-17 : 22 défauts trouvés, dont 15 dans mes propres
correctifs**, sur six tours. Le taux ne baissait pas. Ce n'est pas une formalité.

Commandes de la CI :
```bash
python -m pytest ERP_REACT/backend/tests/ -v --tb=short --timeout=30
cd ERP_REACT/frontend && npx tsc --noEmit && npm run test
```

⚠️ **Mesurer un banc par MUTATION, pas par le nombre de tests verts.** Le
2026-08-17 : 16/16 mutants de la fonction tués, mais **4 mutants du CÂBLAGE
survivaient à 42/42 vert**, dont celui qui réintroduisait le défaut d'origine.

---

## 7. CE QUI RESTE, DANS L'ORDRE

**1. Les courriels en souffrance — le plus rentable, et ça ne dépend d'aucune
infrastructure.** Mesure du 2026-08-17 : **89 fils humains sans réponse**, dont
39 vieux de trois à six mois.

- **Marie-Eve Hermkens** (InnovIA) — apporteur d'affaires, demande une démo
- **Klement Baril** — demande d'information, 17 août
- **Patrick Renaud** (PP Constructions) — 3ᵉ demande pour la même correction
- **Kim Lamontagne** (CPA DT) — comptable, 13 messages, abandonnée depuis le 8 juin
- **Alexy Mathon** (I.C.K.) — « je vous appelle demain », jamais rappelé

→ Charger `profiles/CONSTRUCTO_AI_profil.txt` avant de rédiger. **Brouillon →
montrer → accord → envoyer.** Jamais d'envoi sans accord pour CE message-là.

**2. La défédération GoDaddy** — bloque tout OAuth. Rien à coder avant.

**3. Le rattachement CRM des courriels importés** — `company_id`/`contact_id`
restent NULL alors que le webhook, lui, rattache. C'est la valeur principale du
module. ⚠️ Lire la **vraie** règle dans `emails.py` avant de l'appliquer : la
reproduire de mémoire produirait de faux liens, et un courriel attribué au
mauvais client est pire qu'un courriel non attribué.

**4. Les pièces jointes** — non importées (`has_attachments` forcé à faux
délibérément). Volume utile : ~132 PDF, 12 XLSX, 3 DOCX, 1 DWG. Les 268 PNG sont
des logos de signature, à ne pas importer.

**5. Les quatre autres signatures** — Demo, Robert Cyrenne, Annulation,
Publicité. Même méthode : les extraire d'un envoi qui les porte.

**6. Empaqueter le pont pour les clients** — installateur, **signature de code**
(sans elle SmartScreen et les antivirus le bloquent), mise à jour automatique, et
le **sens inverse** (envoyer depuis l'ERP via Outlook, aujourd'hui absent).

**7. Durcissements OAuth, tous ANTÉRIEURS** — pas de vérification `aud`/`iss`,
state « legacy » à 4 parties sans TTL, HMAC tronqué à 16 hex, aucune clé
étrangère sur `emails.account_id`.

---

## 8. RÉFÉRENCES

| Où | Quoi |
|---|---|
| `CLAUDE.md` (racine du kit) | les règles de conduite, verrouillées dans le code |
| `.claude/skills/courriels-outlook/` | la méthode de tri, de recherche, de rédaction |
| `profiles/CONSTRUCTO_AI_profil.txt` | l'entreprise, l'offre, les 40+ modules, les objections |
| `references/depannage.md` | quand Outlook ne synchronise plus |
| `Constructo_AI_Prod/docs/CONTEXT.md` | leçons **#678-#686** |
| `Constructo_AI_Prod/docs/JOURNAL_SESSIONS.md` | session **S108** |
| `Constructo_AI_Prod/docs/BACKLOG_ORDONNE.md` | section **4quater** |

Identifiants utiles : tenant `tenant_constructo_e1f633`, slug `constructoai`,
compte ERP `id=3` (`info@constructoai.ca`), service Render
`srv-d70vi7dactks738jdn7g`, Postgres `dpg-d4kcvt0gjchc73a3lkk0-a`.

---

*Dernière mise à jour : 2026-08-17. Quand une information s'avère fausse en
usage, la corriger ICI plutôt que de la contourner.*
