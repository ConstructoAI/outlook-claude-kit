"""
outlook_mail.py — Contrôle complet d'une boîte Outlook depuis Python (MAPI/COM).

Lire, chercher, rédiger, répondre, envoyer, classer, supprimer.
Fonctionne avec le profil Outlook déjà authentifié sur le poste : aucun mot de
passe, aucun jeton, aucune inscription d'application, aucun accès administrateur.

Prérequis : Windows, Outlook (classique) installé et LANCÉ, `pip install pywin32`.

    python outlook_mail.py accounts
    python outlook_mail.py folders
    python outlook_mail.py list   --folder inbox --limit 25 [--unread] [--query TEXTE]
    python outlook_mail.py search --query TEXTE [--folder all] [--limit 50]
    python outlook_mail.py read   --id <EntryID> [--html]
    python outlook_mail.py thread --id <EntryID>
    python outlook_mail.py draft  --to a@b.com --subject "..." --body "..." [--cc] [--attach F...]
    python outlook_mail.py reply  --id <EntryID> --body "..." [--all]
    python outlook_mail.py send   --id <EntryID> --yes-send
    python outlook_mail.py move   --id <EntryID> --folder archive
    python outlook_mail.py mark   --id <EntryID> [--unread]
    python outlook_mail.py flag   --id <EntryID> [--off]
    python outlook_mail.py delete --id <EntryID>

Ajoutez --json à n'importe quelle commande pour une sortie exploitable par script.
Ajoutez --account "adresse@domaine" pour cibler une boîte précise (sinon : compte
par défaut du profil).

Deux garde-fous sont codés en dur, volontairement :
  * `send` refuse de partir sans --yes-send.
  * `delete` déplace vers les Éléments supprimés et refuse d'agir sur un message
    qui s'y trouve déjà. Aucune purge définitive n'est exposée.

Licence MIT.
"""
import argparse, json, sys, datetime, os

try:
    import win32com.client as win32
    import pythoncom
except ImportError:
    sys.exit("pywin32 manquant.  ->  pip install pywin32")

# Identifiants de dossiers par défaut d'Outlook (OlDefaultFolders)
OL = dict(deleted=3, outbox=4, sent=5, inbox=6, calendar=9, contacts=10,
          journal=11, notes=12, tasks=13, drafts=16, junk=23)
OL_MAIL_ITEM = 43          # olMail
CREATE_MAIL = 0            # olMailItem


# --------------------------------------------------------------- connexion
def ns():
    return win32.Dispatch("Outlook.Application").GetNamespace("MAPI")


def list_accounts():
    n, out = ns(), []
    for i in range(1, n.Folders.Count + 1):
        try:
            f = n.Folders.Item(i)
            out.append(dict(name=f.Name, items=f.Folders.Count))
        except Exception:
            pass
    return out


def store_root(account=None):
    """Racine du magasin visé. Sans --account : celui de la boîte par défaut."""
    n = ns()
    if account:
        for i in range(1, n.Folders.Count + 1):
            f = n.Folders.Item(i)
            if account.lower() in f.Name.lower():
                return f
        raise SystemExit(f"compte introuvable : {account}\n"
                         f"comptes disponibles : "
                         f"{', '.join(a['name'] for a in list_accounts())}")
    try:
        return n.GetDefaultFolder(OL["inbox"]).Parent
    except Exception:
        return n.Folders.Item(1)


def folder_path(f):
    parts, cur, guard = [], f, 0
    while cur is not None and guard < 20:
        guard += 1
        try:
            parts.append(cur.Name)
            cur = cur.Parent
            if cur is None or not hasattr(cur, "Name"):
                break
        except Exception:
            break
    return "/".join(reversed(parts))


def resolve_folder(name, account=None):
    """Accepte un alias (inbox, sent, drafts…), un nom simple, ou un chemin A/B."""
    name = (name or "inbox").strip()
    key = name.lower()
    if key in OL:
        f = ns().GetDefaultFolder(OL[key])
        if account:                       # forcer le bon magasin
            root = store_root(account)
            if not folder_path(f).startswith(root.Name):
                found = _find(root, [f.Name.lower()])
                if found:
                    return found
        return f
    root = store_root(account)
    if key in ("root", "racine"):
        return root
    parts = [p for p in key.replace("\\", "/").split("/") if p]
    found = _find(root, parts)
    if found is None:
        raise SystemExit(f"dossier introuvable : {name}")
    return found


def _find(root, parts):
    best = None
    for f in _walk(root):
        try:
            fp = folder_path(f).lower()
        except Exception:
            continue
        if fp.endswith("/".join(parts)) or f.Name.lower() == parts[-1]:
            if best is None or len(fp) < len(folder_path(best)):
                best = f
    return best


def _walk(f, depth=0, maxdepth=5):
    yield f
    if depth >= maxdepth:
        return
    try:
        for i in range(1, f.Folders.Count + 1):
            yield from _walk(f.Folders.Item(i), depth + 1, maxdepth)
    except Exception:
        return


# --------------------------------------------------------------- lecture
def _dt(v):
    try:
        return datetime.datetime(v.year, v.month, v.day,
                                 v.hour, v.minute, v.second).isoformat(" ")
    except Exception:
        return None


def brief(m, body_chars=0):
    """Dictionnaire normalisé décrivant un message. body_chars=-1 => corps entier."""
    d = {}
    for attr, key in (("EntryID", "id"), ("Subject", "subject"),
                      ("SenderName", "from"), ("SenderEmailAddress", "from_addr"),
                      ("To", "to"), ("CC", "cc"), ("Categories", "categories"),
                      ("ConversationTopic", "topic")):
        try:
            d[key] = getattr(m, attr)
        except Exception:
            d[key] = None
    for attr, key in (("ReceivedTime", "received"), ("SentOn", "sent")):
        try:
            d[key] = _dt(getattr(m, attr))
        except Exception:
            d[key] = None
    for attr, key, dflt in (("UnRead", "unread", False), ("Size", "size", 0),
                            ("FlagStatus", "flag", 0), ("Importance", "importance", 1)):
        try:
            d[key] = getattr(m, attr)
        except Exception:
            d[key] = dflt
    try:
        d["attachments"] = [a.FileName for a in m.Attachments]
    except Exception:
        d["attachments"] = []
    try:
        d["folder"] = folder_path(m.Parent)
    except Exception:
        d["folder"] = None
    if body_chars:
        try:
            b = " ".join((m.Body or "").split())
            d["body"] = b if body_chars < 0 else b[:body_chars]
        except Exception:
            d["body"] = None
    return d


def iter_items(folder, limit=None, unread=False, newest=True):
    items = folder.Items
    try:
        items.Sort("[ReceivedTime]", newest)
    except Exception:
        pass
    if unread:
        try:
            items = items.Restrict("[UnRead] = True")
            items.Sort("[ReceivedTime]", newest)
        except Exception:
            pass
    out = []
    for it in items:
        try:
            if it.Class != OL_MAIL_ITEM:
                continue
        except Exception:
            continue
        out.append(it)
        if limit and len(out) >= limit:
            break
    return out


def matches(m, q):
    q = q.lower()
    for attr in ("Subject", "SenderName", "SenderEmailAddress", "To", "CC"):
        try:
            v = getattr(m, attr)
            if v and q in str(v).lower():
                return True
        except Exception:
            pass
    try:
        if m.Body and q in m.Body.lower():
            return True
    except Exception:
        pass
    return False


# --------------------------------------------------------------- commandes
def cmd_accounts(a):
    rows = list_accounts()
    print(json.dumps(rows, ensure_ascii=False, indent=2) if a.json
          else "\n".join("  " + r["name"] for r in rows))


def cmd_folders(a):
    rows = []
    for f in _walk(store_root(a.account)):
        try:
            rows.append(dict(path=folder_path(f), items=f.Items.Count,
                             unread=f.UnReadItemCount,
                             depth=folder_path(f).count("/")))
        except Exception:
            pass
    if a.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        base = min((r["depth"] for r in rows), default=0)
        for r in rows:
            if r["items"] or r["unread"]:
                pad = "  " * (r["depth"] - base)
                print(f"{pad}{r['path'].split('/')[-1]:<38} "
                      f"{r['items']:>6} élts  {r['unread']:>4} non lus")


def cmd_list(a):
    f = resolve_folder(a.folder, a.account)
    items = iter_items(f, (a.limit * 6 if a.query else a.limit), a.unread)
    if a.query:
        items = [m for m in items if matches(m, a.query)][:a.limit]
    emit([brief(m, 160) for m in items], a)


def cmd_search(a):
    if a.folder and a.folder.lower() != "all":
        folders = [resolve_folder(a.folder, a.account)]
    else:
        folders = list(_walk(store_root(a.account), maxdepth=3))
    rows = []
    for f in folders:
        try:
            for m in iter_items(f, 500):
                if matches(m, a.query):
                    rows.append(brief(m, 160))
                    if len(rows) >= a.limit:
                        break
        except Exception:
            continue
        if len(rows) >= a.limit:
            break
    emit(rows, a)


def cmd_read(a):
    m = ns().GetItemFromID(a.id)
    d = brief(m, -1)
    if a.html:
        try:
            d["html"] = m.HTMLBody
        except Exception:
            d["html"] = None
    if a.json:
        print(json.dumps(d, ensure_ascii=False, indent=2))
        return
    print(f"De      : {d['from']} <{d['from_addr']}>")
    print(f"À       : {d['to']}")
    if d["cc"]:
        print(f"Cc      : {d['cc']}")
    print(f"Date    : {d['received'] or d['sent']}")
    print(f"Dossier : {d['folder']}")
    print(f"Objet   : {d['subject']}")
    if d["attachments"]:
        print(f"PJ      : {', '.join(d['attachments'])}")
    print("-" * 72)
    print(d.get("body") or "")


def cmd_thread(a):
    src = ns().GetItemFromID(a.id)
    topic = (src.ConversationTopic or src.Subject or "").strip().lower()
    rows = []
    for f in _walk(store_root(a.account), maxdepth=3):
        try:
            for m in iter_items(f, 500):
                if (m.ConversationTopic or "").strip().lower() == topic:
                    rows.append(brief(m, 400))
        except Exception:
            continue
    rows.sort(key=lambda r: r["received"] or r["sent"] or "")
    emit(rows, a)


def cmd_draft(a):
    m = win32.Dispatch("Outlook.Application").CreateItem(CREATE_MAIL)
    m.To = a.to
    if a.cc:
        m.CC = a.cc
    if a.bcc:
        m.BCC = a.bcc
    m.Subject = a.subject or ""
    if a.html_body:
        m.HTMLBody = a.body or ""
    else:
        m.Body = a.body or ""
    for p in (a.attach or []):
        if os.path.exists(p):
            m.Attachments.Add(os.path.abspath(p))
        else:
            print(f"[avertissement] pièce jointe introuvable : {p}", file=sys.stderr)
    m.Save()
    print(json.dumps(dict(status="brouillon créé", id=m.EntryID,
                          subject=m.Subject, to=m.To), ensure_ascii=False, indent=2))


def cmd_reply(a):
    src = ns().GetItemFromID(a.id)
    r = src.ReplyAll() if a.all else src.Reply()
    r.Body = (a.body or "") + "\n\n" + (r.Body or "")
    r.Save()
    print(json.dumps(dict(status="réponse enregistrée en brouillon", id=r.EntryID,
                          subject=r.Subject, to=r.To), ensure_ascii=False, indent=2))


def cmd_send(a):
    if not a.yes_send:
        raise SystemExit(
            "refus : l'envoi exige --yes-send.\n"
            "Ce drapeau matérialise un accord explicite de l'utilisateur pour CET envoi.")
    m = ns().GetItemFromID(a.id)
    to, subj = m.To, m.Subject
    m.Send()
    print(json.dumps(dict(status="envoyé", to=to, subject=subj),
                     ensure_ascii=False, indent=2))


def cmd_move(a):
    m = ns().GetItemFromID(a.id)
    dest = resolve_folder(a.folder, a.account)
    src = folder_path(m.Parent)
    m.Move(dest)
    print(json.dumps(dict(status="déplacé", de=src, vers=folder_path(dest)),
                     ensure_ascii=False, indent=2))


def cmd_mark(a):
    m = ns().GetItemFromID(a.id)
    m.UnRead = bool(a.unread)
    m.Save()
    print(json.dumps(dict(status="marqué", unread=m.UnRead), ensure_ascii=False, indent=2))


def cmd_flag(a):
    m = ns().GetItemFromID(a.id)
    m.FlagStatus = 0 if a.off else 2
    m.Save()
    print(json.dumps(dict(status="indicateur", flag=m.FlagStatus),
                     ensure_ascii=False, indent=2))


def cmd_delete(a):
    """Déplace vers les Éléments supprimés. RÉVERSIBLE. Jamais de purge définitive."""
    m = ns().GetItemFromID(a.id)
    cur = folder_path(m.Parent).lower()
    if any(k in cur for k in ("supprim", "deleted", "gelöscht", "eliminado")):
        raise SystemExit(
            "refus : ce message est déjà dans les Éléments supprimés.\n"
            "La suppression définitive n'est pas exposée par cet outil.\n"
            "Si vous la voulez, faites-la vous-même depuis Outlook.")
    subj = m.Subject
    m.Delete()
    print(json.dumps(dict(status="déplacé vers les Éléments supprimés (réversible)",
                          subject=subj), ensure_ascii=False, indent=2))


def emit(rows, a):
    if a.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))
        return
    for r in rows:
        mark = ("*" if r.get("unread") else " ") + ("@" if r.get("attachments") else " ")
        print(f"\n{mark} {r.get('received') or r.get('sent')}  [{r.get('folder')}]")
        print(f"    De    : {r.get('from')} <{r.get('from_addr')}>")
        print(f"    Objet : {r.get('subject')}")
        if r.get("body"):
            print(f"    Corps : {r['body'][:200]}")
        print(f"    ID    : {r.get('id')}")
    print(f"\n({len(rows)} message(s))")


# --------------------------------------------------------------- CLI
def main():
    p = argparse.ArgumentParser(
        description="Contrôle complet d'une boîte Outlook via MAPI/COM.")
    p.add_argument("--account", default=None,
                   help="adresse ou nom du magasin (défaut : compte principal)")
    p.add_argument("--json", action="store_true", help="sortie JSON")
    sub = p.add_subparsers(dest="cmd", required=True)

    def add(name, fn, help=""):
        s = sub.add_parser(name, help=help)
        s.set_defaults(func=fn)
        return s

    add("accounts", cmd_accounts, "lister les comptes du profil")
    add("folders", cmd_folders, "arborescence des dossiers avec compteurs")
    s = add("list", cmd_list, "lister les messages d'un dossier")
    s.add_argument("--folder", default="inbox")
    s.add_argument("--limit", type=int, default=25)
    s.add_argument("--unread", action="store_true")
    s.add_argument("--query", default=None)
    s = add("search", cmd_search, "rechercher dans plusieurs dossiers")
    s.add_argument("--query", required=True)
    s.add_argument("--folder", default="all")
    s.add_argument("--limit", type=int, default=50)
    s = add("read", cmd_read, "lire un message complet")
    s.add_argument("--id", required=True)
    s.add_argument("--html", action="store_true")
    s = add("thread", cmd_thread, "reconstituer un fil de discussion")
    s.add_argument("--id", required=True)
    s = add("draft", cmd_draft, "créer un brouillon")
    s.add_argument("--to", required=True)
    s.add_argument("--subject", default="")
    s.add_argument("--body", default="")
    s.add_argument("--cc", default="")
    s.add_argument("--bcc", default="")
    s.add_argument("--attach", nargs="*")
    s.add_argument("--html-body", action="store_true")
    s = add("reply", cmd_reply, "préparer une réponse en brouillon")
    s.add_argument("--id", required=True)
    s.add_argument("--body", default="")
    s.add_argument("--all", action="store_true")
    s = add("send", cmd_send, "envoyer un brouillon existant")
    s.add_argument("--id", required=True)
    s.add_argument("--yes-send", action="store_true")
    s = add("move", cmd_move, "déplacer vers un dossier")
    s.add_argument("--id", required=True)
    s.add_argument("--folder", required=True)
    s = add("mark", cmd_mark, "marquer lu / non lu")
    s.add_argument("--id", required=True)
    s.add_argument("--unread", action="store_true")
    s = add("flag", cmd_flag, "poser / retirer l'indicateur de suivi")
    s.add_argument("--id", required=True)
    s.add_argument("--off", action="store_true")
    s = add("delete", cmd_delete, "déplacer vers les Éléments supprimés")
    s.add_argument("--id", required=True)

    a = p.parse_args()
    pythoncom.CoInitialize()
    try:
        a.func(a)
    except Exception as e:
        sys.exit(f"erreur : {e}\n"
                 f"(Outlook est-il lancé ? Le compte existe-t-il ? "
                 f"Essayez : python outlook_mail.py accounts)")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    main()
