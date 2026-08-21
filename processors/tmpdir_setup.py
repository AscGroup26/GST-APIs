"""
Force every temporary file this app writes onto the project drive.

Windows puts %TEMP% on C:. When C: fills up, Excel export dies with

    xlsxwriter.exceptions.FileCreateError: [Errno 28] No space left on device

even though the project drive has hundreds of GB free. Naming the output file
on D: is not enough on its own: xlsxwriter streams every worksheet through its
own temp files taken from tempfile.gettempdir(), and pandas.ExcelWriter does
the same, so the bytes still land on the full drive.

Importing this module first redirects tempfile AND the TMP/TEMP environment
variables, which covers xlsxwriter, pandas and any other library that asks the
OS for scratch space.
"""

import os
import tempfile

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TMP_DIR = os.path.join(PROJECT_ROOT, ".xlsx_tmp")


def use_project_tmpdir(path=None):
    """Point tempfile and TMP/TEMP at the project drive. Returns the path used."""
    target = path or TMP_DIR
    try:
        os.makedirs(target, exist_ok=True)
    except OSError:
        # Can't create it — leave the OS default rather than break startup.
        return tempfile.gettempdir()
    tempfile.tempdir = target
    os.environ["TMP"] = target
    os.environ["TEMP"] = target
    return target


def purge(older_than_seconds=None):
    """
    Delete leftover files in the scratch dir.

    Crashed exports leave .xlsx files behind; without this they accumulate on
    the project drive until it hits the same wall C: did.
    """
    import time

    removed = 0
    if not os.path.isdir(TMP_DIR):
        return removed
    now = time.time()
    for name in os.listdir(TMP_DIR):
        full = os.path.join(TMP_DIR, name)
        if not os.path.isfile(full):
            continue
        if older_than_seconds is not None:
            try:
                if now - os.path.getmtime(full) < older_than_seconds:
                    continue
            except OSError:
                continue
        try:
            os.unlink(full)
            removed += 1
        except OSError:
            pass
    return removed
