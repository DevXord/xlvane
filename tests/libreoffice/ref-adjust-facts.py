# Observes how LibreOffice Calc adjusts references when rows or columns are inserted or deleted
# (docs/facts/ref-adjust.md). Usage:
#   python3 ref-adjust-facts.py <soffice> <work-dir> <jobs.json>
# jobs.json: [{"input": "rows.xlsx", "output": "rows-insert.xlsx", "sheet": "Main",
#              "axis": "row" | "col", "op": "insert" | "delete", "at": 5, "count": 2}, ...]
# "at" is 1-based (row number or column number). Each job loads its input afresh, applies one
# operation through the UNO API (XTableRows/XTableColumns insertByIndex/removeByIndex) and saves
# the result with the "Calc Office Open XML" filter. Needs LibreOffice's Python UNO bridge.
import json
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def main():
    soffice, work, jobs_path = sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3]
    with open(jobs_path, encoding='utf-8') as f:
        jobs = json.load(f)
    profile = 'file://' + os.path.join(work, 'lo-profile')
    pipe = 'xlvane_refadj_%d' % os.getpid()
    proc = subprocess.Popen([soffice, '--headless', '--invisible', '--norestore', '--nologo',
                             '-env:UserInstallation=' + profile,
                             '--accept=pipe,name=%s;urp;' % pipe],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        ctx = None
        for _ in range(240):
            try:
                ctx = resolver.resolve('uno:pipe,name=%s;urp;StarOffice.ComponentContext' % pipe)
                break
            except Exception:
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError('could not connect to LibreOffice')
        desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        for job in jobs:
            url = uno.systemPathToFileUrl(os.path.join(work, job['input']))
            doc = desktop.loadComponentFromURL(url, '_blank', 0, (prop('Hidden', True),))
            sheet = doc.Sheets.getByName(job['sheet'])
            coll = sheet.Rows if job['axis'] == 'row' else sheet.Columns
            if job['op'] == 'insert':
                coll.insertByIndex(job['at'] - 1, job['count'])
            else:
                coll.removeByIndex(job['at'] - 1, job['count'])
            out = uno.systemPathToFileUrl(os.path.join(work, job['output']))
            doc.storeToURL(out, (prop('FilterName', 'Calc Office Open XML'),))
            doc.close(True)
        try:
            desktop.terminate()
        except Exception:
            pass
    finally:
        try:
            proc.wait(timeout=30)
        except Exception:
            proc.kill()


if __name__ == '__main__':
    main()
