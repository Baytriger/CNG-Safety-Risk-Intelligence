# -*- coding: utf-8 -*-
"""Headless LibreOffice conversion that forces a full recalculation of OOXML formulas on load."""
import os, subprocess, tempfile

XCU = ('<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry" '
       'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
       '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>'
       '</oor:items>')


def recalc_convert(src, outdir, fmt="xlsx:Calc MS Excel 2007 XML"):
    prof = tempfile.mkdtemp(prefix="lo_profile_")
    os.makedirs(os.path.join(prof, "user"), exist_ok=True)
    open(os.path.join(prof, "user", "registrymodifications.xcu"), "w").write(XCU)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", fmt, "--outdir", outdir, src],
                   capture_output=True, timeout=300)
