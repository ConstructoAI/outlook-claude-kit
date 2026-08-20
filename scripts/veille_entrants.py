# -*- coding: utf-8 -*-
"""Veille sur les courriels ENTRANTS. Sort des qu'un nouveau message arrive.

    python scripts/veille_entrants.py [--minutes 480] [--intervalle 60]

POURQUOI DEUX SIGNAUX, ET PAS UN SEUL.

Le fichier `.ost` est le magasin hors ligne d'Outlook. Sa date de modification
bouge des qu'Outlook synchronise, ce qui en fait un declencheur tres bon marche :
le lire coute un `stat`, la ou interroger MAPI coute une connexion COM.

Mais il est BRUYANT. Outlook y ecrit pour des dizaines de raisons qui ne sont pas
du courrier entrant : un message marque lu, un rendez-vous, un dossier reindexe,
une purge de cache. Annoncer un courriel sur la seule foi du `.ost` reviendrait a
crier au loup plusieurs fois par heure.

D'ou la regle : le `.ost` DECLENCHE la mesure, il ne la REMPLACE pas. Un nouveau
message se reconnait a une date de reception plus recente que la derniere vue --
jamais au fait que le fichier a grossi.

Sort avec le code 0 des qu'un entrant est trouve, en le decrivant. Code 1 si le
delai s'epuise sans rien voir.
"""
import argparse
import glob
import io
import json
import os
import subprocess
import sys
import time

# `line_buffering` n'est pas un detail : hors terminal, Python retient sa sortie
# jusqu'a la fin du programme. Une veille de huit heures resterait donc MUETTE,
# et on la croirait morte alors qu'elle travaille.
sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

ICI = os.path.dirname(os.path.abspath(__file__))
OUTIL = os.path.join(ICI, "outlook_mail.py")
OST = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Microsoft", "Outlook", "*.ost")


def dernieres(n=8):
    """Les n derniers recus, par le meme outil que tout le reste."""
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    r = subprocess.run([sys.executable, OUTIL, "--json", "list", "--limit", str(n)],
                       capture_output=True, text=True, encoding="utf-8", env=env)
    if r.returncode != 0:
        return None                      # mesure ratee : on ne conclut RIEN
    try:
        return json.loads(r.stdout)
    except Exception:
        return None


def empreinte_ost():
    """{chemin: mtime} des magasins hors ligne. Vide si aucun n'est lisible."""
    e = {}
    for p in glob.glob(OST):
        try:
            e[p] = os.path.getmtime(p)
        except OSError:
            pass
    return e


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--minutes", type=int, default=480)
    ap.add_argument("--intervalle", type=int, default=60)
    a = ap.parse_args()

    base = dernieres()
    if base is None:
        sys.exit("ARRET : impossible de lire la boite au demarrage. Outlook est-il lance ?")
    repere = max((m.get("received") or "") for m in base) if base else ""
    connus = {m.get("id") for m in base}
    ost = empreinte_ost()

    print("veille demarree — repere : %s" % (repere or "(boite vide)"))
    print("magasins surveilles : %d" % len(ost))
    if not ost:
        print("AVERTISSEMENT : aucun .ost trouve, la veille interrogera MAPI a chaque tour.")

    tours = max(1, (a.minutes * 60) // a.intervalle)
    for i in range(1, tours + 1):
        time.sleep(a.intervalle)

        nouveau_ost = empreinte_ost()
        bouge = (nouveau_ost != ost)
        ost = nouveau_ost

        # Sans .ost lisible, on ne peut pas pre-filtrer : on mesure a chaque tour.
        if ost and not bouge:
            continue

        msgs = dernieres()
        if msgs is None:
            print("[tour %d] MESURE RATEE (on ne conclut rien)" % i)
            continue

        neufs = [m for m in msgs
                 if (m.get("received") or "") > repere and m.get("id") not in connus]
        if not neufs:
            continue

        neufs.sort(key=lambda m: m.get("received") or "")
        print()
        print("=== %d COURRIEL(S) ENTRANT(S) ===" % len(neufs))
        for m in neufs:
            print("  %s  de %s <%s>" % (m.get("received"), m.get("from"), m.get("from_addr")))
            print("     objet : %s" % m.get("subject"))
            pj = m.get("attachments") or []
            if pj:
                print("     PJ    : %s" % ", ".join(pj))
            print("     id    : %s" % m.get("id"))
        return 0

    print("delai epuise (%d min) — aucun entrant." % a.minutes)
    return 1


if __name__ == "__main__":
    sys.exit(main())
