"""
embed_taskpane.py – Setter inn Office.js-oppgavepanel direkte i matte.docx.
Når dokumentet åpnes i Word, vises Notater-panelet automatisk på høyre side.
"""

import zipfile
import shutil
import os
import winreg
from pathlib import Path
from utils import find_matte_docx

MANIFEST_ID = "a1b2c3d4-e5f6-7890-abcd-ef1234567890"

WEBEXT_XML = f"""\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension
  xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  id="{{{MANIFEST_ID}}}">
  <we:reference id="{MANIFEST_ID}" version="1.0.0.0"
    store="__ADDIN_URL__" storeType="FileSystem"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="taskpaneId" value="NotaterPanel"/>
    <we:property name="version" value="1.0.0.0"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
"""

TASKPANE_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<tp:taskpanes xmlns:tp="http://schemas.microsoft.com/office/drawing/2010/main">
  <tp:taskpane dockstate="msoPaneStateDocked" visibility="1" width="350" row="4">
    <tp:webextensionref
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      r:id="rId_we1"/>
  </tp:taskpane>
</tp:taskpanes>
"""


def _add_trusted_catalog(addin_dir: Path):
    """Legger addin-mappen til som klarert filkatalog i Word."""
    # Word forventer file:/// med forward slashes, uten %-koding av mellomrom
    url = "file:///" + str(addin_dir).replace("\\", "/")  # file:///C:/path/with spaces/...
    guid = "{A1B2C3D4-0000-0000-0000-NOTATER00001}"
    try:
        key = winreg.CreateKey(
            winreg.HKEY_CURRENT_USER,
            rf"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{guid}"
        )
        winreg.SetValueEx(key, "Url",   0, winreg.REG_SZ,    url)
        winreg.SetValueEx(key, "Flags", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        print(f"[Addin] Klarert katalog satt: {url}")
    except Exception as e:
        print(f"[Addin] Registerfeil: {e}")


def embed_into_docx(doc_path: Path, addin_url: str = ""):
    """Legger til webextension og taskpane-deler i matte.docx."""
    tmp = doc_path.with_suffix(".tmp.docx")

    with zipfile.ZipFile(str(doc_path), "r") as zin, \
         zipfile.ZipFile(str(tmp), "w", zipfile.ZIP_DEFLATED) as zout:

        names = zin.namelist()

        # Kopier alle eksisterende filer, oppdater de vi trenger
        for name in names:
            data = zin.read(name)

            if name == "[Content_Types].xml":
                txt = data.decode("utf-8")
                # Legg til content types hvis de ikke finnes
                if "webextension" not in txt:
                    txt = txt.replace(
                        "</Types>",
                        '<Override PartName="/word/webextensions/webextension1.xml"'
                        ' ContentType="application/vnd.ms-office.webextension+xml"/>\n'
                        '<Override PartName="/word/taskpanes/taskpane1.xml"'
                        ' ContentType="application/vnd.ms-office.activeX+xml"/>\n'
                        "</Types>"
                    )
                zout.writestr(name, txt.encode("utf-8"))

            elif name == "word/_rels/document.xml.rels":
                txt = data.decode("utf-8")
                # Legg til relasjoner for webextension og taskpane
                if "webextension" not in txt:
                    txt = txt.replace(
                        "</Relationships>",
                        '<Relationship Id="rId_we1"'
                        ' Type="http://schemas.microsoft.com/office/2011/relationships/webextension"'
                        ' Target="webextensions/webextension1.xml"/>\n'
                        '<Relationship Id="rId_tp1"'
                        ' Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"'
                        ' Target="taskpanes/taskpane1.xml"/>\n'
                        "</Relationships>"
                    )
                zout.writestr(name, txt.encode("utf-8"))

            else:
                zout.writestr(name, data)

        # Legg til nye deler
        if "word/webextensions/webextension1.xml" not in names:
            xml = WEBEXT_XML.replace("__ADDIN_URL__", addin_url)
            zout.writestr("word/webextensions/webextension1.xml",
                          xml.encode("utf-8"))
        if "word/taskpanes/taskpane1.xml" not in names:
            zout.writestr("word/taskpanes/taskpane1.xml",
                          TASKPANE_XML.encode("utf-8"))

    # Erstatt originalen
    os.replace(str(tmp), str(doc_path))
    print(f"[Addin] Oppgavepanel innebygd i {doc_path.name}")


def run():
    addin_dir = Path(__file__).parent.parent / "addin"
    _add_trusted_catalog(addin_dir)

    addin_url = "file:///" + str(addin_dir).replace("\\", "/")

    doc_path = find_matte_docx()
    if not doc_path:
        print("[Addin] Finner ikke matte.docx – hopper over dokumentendring")
        return

    embed_into_docx(doc_path, addin_url)
    print("[Addin] Ferdig! Åpne matte.docx på nytt for å se panelet.")


if __name__ == "__main__":
    run()
