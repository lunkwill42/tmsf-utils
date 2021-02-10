#!/usr/bin/env python3
# coding: utf-8
"""
Sorterer listen over aktive medlemmer i TMSF etter stemmegrupper og åpner
en tom e-post til hver gruppe.
"""

from collections import namedtuple
import os.path
import webbrowser
from urllib.parse import quote

import xlrd

DROPBOXFIL = "Trøndernes MSF/Diverse/Medlemsliste/Adresseliste aktive og passive.xls"
STEMMEGRUPPER = ("1.tenor", "2.tenor", "1.bass", "2.bass")


def main():
    aktive = [
        medlem for medlem in hent_medlemsliste() if "aktiv" in medlem.aktiv.lower()
    ]
    for gruppe in STEMMEGRUPPER:
        mailto(gruppe, [medlem for medlem in aktive if gruppe in medlem.stemme.lower()])


def hent_medlemsliste(filnavn=None):
    if not filnavn:
        filnavn = finn_listefil()
    assert filnavn, "Fant ikke regnearket med medlemsregisteret"
    bok = xlrd.open_workbook(filnavn)
    medlemsark = bok.sheets()[0]
    Medlem = namedtuple(
        "Medlem",
        [medlemsark.cell(0, x).value.lower() for x in range(0, medlemsark.ncols)],
    )
    return [
        Medlem(*[medlemsark.cell(row, col).value for col in range(medlemsark.ncols)])
        for row in range(1, medlemsark.nrows)
    ]


def finn_listefil():
    for prefix in ("~/Dropbox", "~/privat/Dropbox"):
        filename = os.path.join(os.path.expanduser(prefix), DROPBOXFIL)
        if os.path.exists(filename):
            return filename


def mailto(subject, medlemmer):
    rcpt = quote(
        ",".join(f'"{m.fornavn} {m.etternavn}" <{m.epost}>' for m in medlemmer)
    )
    subject = quote(subject)
    mailto = f"mailto:{rcpt}?subject={subject}"
    webbrowser.open(mailto)


if __name__ == "__main__":
    main()
