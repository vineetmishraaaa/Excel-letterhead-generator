import os
import sys
import time
import socket
import shutil
import tempfile
import subprocess
from pathlib import Path

# The Debian python3-uno package installs UNO under this path.
sys.path.append('/usr/lib/python3/dist-packages')

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    item = PropertyValue()
    item.Name = name
    item.Value = value
    return item


def wait_for_port(host, port, timeout=30):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection((host, port), timeout=0.5):
                return True
        except OSError:
            time.sleep(0.25)
    return False


def get_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


def export_workbook(input_path, output_path):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve()

    if not input_path.exists():
        raise FileNotFoundError(f'Excel file not found: {input_path}')

    port = get_free_port()
    profile_dir = Path(tempfile.mkdtemp(prefix='lo-profile-'))
    libreoffice = None
    doc = None

    try:
        profile_url = uno.systemPathToFileUrl(str(profile_dir))

        cmd = [
            'libreoffice',
            '--headless',
            '--nologo',
            '--nodefault',
            '--nofirststartwizard',
            '--norestore',
            f'-env:UserInstallation={profile_url}',
            f'--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext',
        ]

        libreoffice = subprocess.Popen(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        if not wait_for_port('127.0.0.1', port, timeout=30):
            raise RuntimeError('LibreOffice did not start.')

        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            'com.sun.star.bridge.UnoUrlResolver',
            local_ctx,
        )

        context = resolver.resolve(
            f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext'
        )

        service_manager = context.ServiceManager
        desktop = service_manager.createInstanceWithContext(
            'com.sun.star.frame.Desktop',
            context,
        )

        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(input_path)),
            '_blank',
            0,
            (
                prop('Hidden', True),
                prop('ReadOnly', True),
                prop('UpdateDocMode', 0),
            ),
        )

        if doc is None:
            raise RuntimeError('LibreOffice could not open the Excel workbook.')

        # Keep the important layout behavior from the proven Windows script.
        page_styles = doc.StyleFamilies.getByName('PageStyles')

        visible_count = 0

        for index in range(doc.Sheets.getCount()):
            sheet = doc.Sheets.getByIndex(index)

            if not sheet.IsVisible:
                continue

            visible_count += 1

            try:
                page_style = page_styles.getByName(sheet.PageStyle)

                # A4 portrait, matching the target output.
                page_style.Width = 21000
                page_style.Height = 29700
                page_style.IsLandscape = False

                # Fit horizontally to one page, allow vertical overflow.
                page_style.ScaleToPages = 1
                page_style.ScaleToPagesX = 1
                page_style.ScaleToPagesY = 0

                # Similar margins to the original working workflow.
                page_style.LeftMargin = 1000
                page_style.RightMargin = 1000
                page_style.TopMargin = 500
                page_style.BottomMargin = 500

                page_style.CenterHorizontally = True
                page_style.CenterVertically = False
                page_style.PrintGrid = False
                page_style.PrintHeaders = False
                page_style.PrintAnnotations = False

            except Exception as exc:
                raise RuntimeError(
                    f'Could not configure worksheet "{sheet.Name}": {exc}'
                ) from exc

        if visible_count == 0:
            raise RuntimeError('The workbook has no visible worksheets.')

        output_path.parent.mkdir(parents=True, exist_ok=True)

        pdf_url = uno.systemPathToFileUrl(str(output_path))

        doc.storeToURL(
            pdf_url,
            (
                prop('FilterName', 'calc_pdf_Export'),
                prop('Overwrite', True),
            ),
        )

        if not output_path.exists():
            raise RuntimeError('LibreOffice did not create the PDF.')

    finally:
        if doc is not None:
            try:
                doc.close(True)
            except Exception:
                pass

        if libreoffice is not None:
            try:
                libreoffice.terminate()
                libreoffice.wait(timeout=5)
            except Exception:
                try:
                    libreoffice.kill()
                except Exception:
                    pass

        shutil.rmtree(profile_dir, ignore_errors=True)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Usage: python libreoffice_export.py INPUT.xlsx OUTPUT.pdf')
        sys.exit(2)

    try:
        export_workbook(sys.argv[1], sys.argv[2])
    except Exception as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        sys.exit(1)
