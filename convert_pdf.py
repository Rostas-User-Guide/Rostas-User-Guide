#!/usr/bin/env python3
"""
Converts Rostas_Coordinator_Guide.docx to PDF via LibreOffice UNO.
Updates all fields and TOC indexes before exporting so page numbers are correct.
Usage: python3 convert_pdf.py
"""
import subprocess, time, os, sys, threading

TIMEOUT_SECONDS = 300  # hard ceiling for the whole conversion — 5 minutes

def watchdog(proc):
    """If we're still running after TIMEOUT_SECONDS, kill LibreOffice and exit hard."""
    time.sleep(TIMEOUT_SECONDS)
    print(f'✗ TIMEOUT after {TIMEOUT_SECONDS}s — killing LibreOffice and failing the job.')
    try:
        proc.kill()
    except Exception:
        pass
    os._exit(1)

def main():
    cwd  = os.getcwd()
    docx = os.path.join(cwd, 'Rostas_Coordinator_Guide.docx')
    pdf  = os.path.join(cwd, 'Rostas_Coordinator_Guide.pdf')
    docx_url = f'file://{docx}'
    pdf_url  = f'file://{pdf}'
    profile_url = 'file:///tmp/lo_profile_' + str(os.getpid())

    print('Starting LibreOffice listener...')
    proc = subprocess.Popen([
        'libreoffice', '--headless', '--norestore', '--nofirststartwizard',
        '--nologo', '--nolockcheck', '--nodefault',
        f'-env:UserInstallation={profile_url}',
        '--accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManager'
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # Watchdog runs in the background for the whole script; kills everything if we hang.
    watchdog_thread = threading.Thread(target=watchdog, args=(proc,), daemon=True)
    watchdog_thread.start()

    try:
        import uno
        from com.sun.star.beans import PropertyValue

        localCtx  = uno.getComponentContext()
        localSmgr = localCtx.ServiceManager
        resolver  = localSmgr.createInstanceWithContext(
            'com.sun.star.bridge.UnoUrlResolver', localCtx)

        print('Connecting to LibreOffice...')
        ctx = None
        last_err = None
        for attempt in range(30):  # retry for up to ~30s instead of one fixed 8s guess
            try:
                ctx = resolver.resolve(
                    'uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
                break
            except Exception as e:
                last_err = e
                time.sleep(1)
        if ctx is None:
            raise RuntimeError(f'Could not connect to LibreOffice after 30s: {last_err}')

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)

        print(f'Opening {os.path.basename(docx)}...')
        # MacroExecutionMode=0 (NEVER_EXECUTE) avoids any macro-security prompt in headless mode.
        macro_prop = PropertyValue()
        macro_prop.Name = 'MacroExecutionMode'
        macro_prop.Value = 0
        doc = desktop.loadComponentFromURL(docx_url, '_blank', 0, (macro_prop,))

        print('Updating TOC and all fields...')
        doc.getTextFields().refresh()
        dispatcher = smgr.createInstanceWithContext(
            'com.sun.star.frame.DispatchHelper', ctx)
        frame = doc.getCurrentController().Frame
        dispatcher.executeDispatch(frame, '.uno:UpdateAllIndexes', '', 0, ())
        dispatcher.executeDispatch(frame, '.uno:UpdateFields',     '', 0, ())

        print('Exporting PDF...')
        p = PropertyValue()
        p.Name  = 'FilterName'
        p.Value = 'writer_pdf_Export'
        doc.storeToURL(pdf_url, (p,))
        doc.close(True)

        size = os.path.getsize(pdf) / 1024 / 1024
        print(f'✓ PDF created: {os.path.basename(pdf)} ({size:.1f} MB)')

    except Exception as e:
        print(f'✗ Conversion failed: {e}')
        proc.terminate()
        sys.exit(1)

    proc.terminate()

main()
