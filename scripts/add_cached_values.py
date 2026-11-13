"""Calculates every formula and stores the result in the file so viewers that
don't recalculate (previews, phones) still show numbers. Excel recalculates on open."""
import sys, zipfile, shutil, re
import formulas, numpy as np
from lxml import etree

path = sys.argv[1]
sol = formulas.ExcelModel().loads(path).finish().calculate()
vals = {}
for k, v in sol.items():
    m = re.match(r"'\[[^\]]+\]([^']+)'!([A-Z]+[0-9]+)$", k)
    if m:
        vals[(m.group(1).upper(), m.group(2))] = v.value[0, 0]

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
tmp = path + ".tmp"
with zipfile.ZipFile(path) as zin:
    wbx = etree.fromstring(zin.read("xl/workbook.xml"))
    rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels}
    sheetfile = {}
    for s in wbx.find(f"{{{NS}}}sheets"):
        t = target[s.get(f"{{{RNS}}}id")].lstrip("/")
        sheetfile[t if t.startswith("xl/") else "xl/" + t] = s.get("name").upper()
    n = 0
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in sheetfile:
                root = etree.fromstring(data)
                for c in root.iter(f"{{{NS}}}c"):
                    f = c.find(f"{{{NS}}}f")
                    if f is None:
                        continue
                    val = vals[(sheetfile[item.filename], c.get("r"))]
                    v = c.find(f"{{{NS}}}v")
                    if v is None:
                        v = etree.SubElement(c, f"{{{NS}}}v")
                    if isinstance(val, (str, np.str_)):
                        c.set("t", "str"); v.text = str(val)
                    elif isinstance(val, (bool, np.bool_)):
                        c.set("t", "b"); v.text = "1" if val else "0"
                    else:
                        fv = float(val)
                        if "t" in c.attrib: del c.attrib["t"]
                        v.text = repr(int(fv)) if fv == int(fv) else repr(fv)
                    n += 1
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
shutil.move(tmp, path)
print("cached values written for", n, "formula cells")
