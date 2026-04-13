#!/usr/bin/env python3
"""
Converts Rostas_Coordinator_Guide.docx to PDF via LibreOffice UNO.
Updates all fields and TOC indexes before exporting so page numbers are correct.

Usage: python3 convert_pdf.py
"""
import subprocess, time, os, sys

def main():
    cwd  = os.getcwd()
    docx = os.path.join(cwd, 'Rostas_Coordinator_Guide.docx')
    pdf  = os.path.join(cwd, 'Rostas_Coordinator_Guide.pdf')

    docx_url = f'file://{docx}'
    pdf_url  = f'file://{pdf}'

    print('Starting LibreOffice listener...')
    proc = subprocess.Popen([
        'libreoffice', '--headless', '--norestore', '--nofirststartwizard',
        '--accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManager'
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(8)  # Wait for LO to be ready

    import uno
    from com.sun.star.beans import PropertyValue

    localCtx  = uno.getComponentContext()
    localSmgr = localCtx.ServiceManager
    resolver  = localSmgr.createInstanceWithContext(
        'com.sun.star.bridge.UnoUrlResolver', localCtx)

    print('Connecting to LibreOffice...')
    ctx  = resolver.resolve(
        'uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext')
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)

    print(f'Opening {os.path.basename(docx)}...')
    doc = desktop.loadComponentFromURL(docx_url, '_blank', 0, ())

    print('Updating TOC and all fields...')
    doc.getTextFields().refresh()
    dispatcher = smgr.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    frame = doc.getCurrentController().Frame
    dispatcher.executeDispatch(frame, '.uno:UpdateAllIndexes', '', 0, ())
    dispatcher.executeDispatch(frame, '.uno:UpdateFields',     '', 0, ())

    print(f'Exporting PDF...')
    p = PropertyValue()
    p.Name  = 'FilterName'
    p.Value = 'writer_pdf_Export'
    doc.storeToURL(pdf_url, (p,))
    doc.close(True)
    proc.terminate()

    size = os.path.getsize(pdf) / 1024 / 1024
    print(f'✓ PDF created: {os.path.basename(pdf)} ({size:.1f} MB)')

main()
