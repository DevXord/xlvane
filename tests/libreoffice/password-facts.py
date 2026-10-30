# Observes LibreOffice Calc and sheet passwords (docs/facts/sheet-rules.md SR-55, SR-58). Usage:
#   python3 password-facts.py <soffice> <work-dir> <password> [<file-to-check>]
# Without <file-to-check>: creates a spreadsheet, protects sheet 1 and the document with
# <password> through the UNO API (XProtectable.protect) and saves <work-dir>/protected.xlsx with
# the "Calc Office Open XML" filter.
# With <file-to-check> (in <work-dir>): opens it, tries XProtectable.unprotect on sheet 1 with a
# wrong password and then with <password>, and prints a JSON line
# {"afterWrong": <isProtected>, "afterRight": <isProtected>}.
# Needs LibreOffice's Python UNO bridge.
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
    soffice, work, password = sys.argv[1], os.path.abspath(sys.argv[2]), sys.argv[3]
    profile = 'file://' + os.path.join(work, 'lo-profile')
    pipe = 'xlvane_pw_%d' % os.getpid()
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
        if len(sys.argv) > 4:
            url = uno.systemPathToFileUrl(os.path.join(work, sys.argv[4]))
            doc = desktop.loadComponentFromURL(url, '_blank', 0, (prop('Hidden', True),))
            sheet = doc.Sheets.getByIndex(0)
            result = {}
            for key, attempt in (('afterWrong', password + '-wrong'), ('afterRight', password)):
                try:
                    sheet.unprotect(attempt)
                except Exception:
                    pass
                result[key] = bool(sheet.isProtected())
            print(json.dumps(result))
            doc.close(True)
        else:
            doc = desktop.loadComponentFromURL('private:factory/scalc', '_blank', 0, (prop('Hidden', True),))
            sheet = doc.Sheets.getByIndex(0)
            sheet.getCellByPosition(0, 0).setString('protected')
            sheet.protect(password)
            doc.protect(password)
            out = uno.systemPathToFileUrl(os.path.join(work, 'protected.xlsx'))
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
