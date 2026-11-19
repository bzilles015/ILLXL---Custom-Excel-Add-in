"""Add the ILLXL ribbon (src/ribbon.xml + assets/icons/*.png) to a built .xlam.

An .xlam is a zip package. Excel can't add a ribbon itself, so after build.ps1
has saved the add-in this script rewrites the package with:
  customUI/customUI14.xml                 the ribbon definition
  customUI/images/<name>.png              every icon the XML refers to (image="...")
  customUI/_rels/customUI14.xml.rels      links image ids -> png files
  _rels/.rels                             tells Excel the package has a ribbon
  [Content_Types].xml                     registers .png
Usage:  python -I tools/inject_ribbon.py <addin.xlam> <ribbon.xml> <icons dir>
"""
import re
import sys
import zipfile
from pathlib import Path
from xml.dom import minidom

UI_REL = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
IMG_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def main(xlam: Path, ribbon: Path, icons: Path) -> None:
    xml = ribbon.read_text(encoding="utf-8")
    minidom.parseString(xml.encode("utf-8"))            # fail the build on malformed XML
    code = re.sub(r"<!--.*?-->", "", xml, flags=re.S)  # ignore examples in comments
    images = sorted(set(re.findall(r'\bimage="([^"]+)"', code)))
    missing = [i for i in images if not (icons / f"{i}.png").exists()]
    if missing:
        sys.exit(f"ribbon.xml refers to icons that don't exist: {missing}")

    with zipfile.ZipFile(xlam) as z:
        parts = {n: z.read(n) for n in z.namelist()}

    # drop any earlier ribbon (rebuilds are repeatable)
    parts = {n: b for n, b in parts.items() if not n.startswith("customUI/")}

    rels = parts["_rels/.rels"].decode("utf-8")
    rels = re.sub(r'<Relationship [^>]*ui/extensibility"[^>]*/>', "", rels)
    rels = rels.replace("</Relationships>",
                        f'<Relationship Id="rIdILLXLRibbon" Type="{UI_REL}" Target="customUI/customUI14.xml"/></Relationships>')
    parts["_rels/.rels"] = rels.encode("utf-8")

    ct = parts["[Content_Types].xml"].decode("utf-8")
    if 'Extension="png"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="png" ContentType="image/png"/><Default ', 1)
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    parts["customUI/customUI14.xml"] = xml.encode("utf-8")
    img_rels = "".join(f'<Relationship Id="{i}" Type="{IMG_REL}" Target="images/{i}.png"/>' for i in images)
    parts["customUI/_rels/customUI14.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{img_rels}</Relationships>").encode("utf-8")
    for i in images:
        parts[f"customUI/images/{i}.png"] = (icons / f"{i}.png").read_bytes()

    tmp = xlam.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))   # must come first
        for n, b in parts.items():
            z.writestr(n, b)
    tmp.replace(xlam)
    print(f"ribbon added: {len(images)} icons")


if __name__ == "__main__":
    main(Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3]))
