"""Make crates/ras-workbook/src/view/handles.zip: the adjust handles of the
preset shapes, for the grid's 調整ハンドル (the yellow handles that change
the adjust values of a shape, `a:avLst`).

The definitions come from presetShapeDefinitions.xml (ECMA-376 Part 1, the
annex "Preset Shape Definitions"; LibreOffice has a copy at
oox/source/drawingml/customshapes/presetShapeDefinitions.xml). Of each
preset with handles (`ahLst`) this keeps what the page evaluates them with:
the adjust values and their defaults (avLst), the guides (gdLst), the
handles (ahLst) and the paths (pathLst, for the outline shown while a handle
is dragged), as JSON in handles.json in the zip (deflated: the crate embeds
the zip and reads it on first use):

    {"roundRect": {"av": [["adj", "val 16667"]],
                   "gd": [["a", "pin 0 adj 50000"], ...],
                   "ah": [{"x": ["adj", "0", "50000"], "pos": ["x1", "t"]}],
                   "path": [{"c": [["M", "l", "x1"], ["A", "x1", "x1", "cd2", "cd4"], ...]}]}}

A handle moves `x`/`y` (ahXY: the guide, its least and greatest value) or
`r`/`a` (ahPolar: radius and angle); a path has `w`/`h` (its own space),
`fill` ("none" when not filled) and `stroke` (false when not stroked); its
commands are M (moveTo), L (lnTo), A (arcTo: wR hR stAng swAng), Q
(quadBezTo), C (cubicBezTo) and Z (close).

    python3 crates/ras-workbook/tools/make_handles.py presetShapeDefinitions.xml

The zip is the same for the same definitions (fixed date, sorted keys).
"""
import json
import pathlib
import sys
import xml.etree.ElementTree as ET
import zipfile

OUT = pathlib.Path(__file__).resolve().parents[1] / "src" / "view" / "handles.zip"


def local(tag):
    return tag.rsplit("}", 1)[-1]


def kid(e, name):
    return next((c for c in e if local(c.tag) == name), None)


def guides(e):
    return [[g.get("name"), g.get("fmla")] for g in e] if e is not None else []


def handle(h):
    out = {}
    if local(h.tag) == "ahXY":
        keys = [("x", "gdRefX", "minX", "maxX"), ("y", "gdRefY", "minY", "maxY")]
    else:
        keys = [("r", "gdRefR", "minR", "maxR"), ("a", "gdRefAng", "minAng", "maxAng")]
    for k, ref, lo, hi in keys:
        if h.get(ref):
            out[k] = [h.get(ref), h.get(lo, "0"), h.get(hi, "0")]
    pos = kid(h, "pos")
    out["pos"] = [pos.get("x"), pos.get("y")]
    return out


def pt(e):
    p = kid(e, "pt")
    return [p.get("x"), p.get("y")]


def path(p):
    out = {}
    for k in ("w", "h"):
        if p.get(k):
            out[k] = int(p.get(k))
    if p.get("fill") == "none":
        out["fill"] = "none"
    if p.get("stroke") in ("0", "false"):
        out["stroke"] = False
    cmds = []
    for c in p:
        n = local(c.tag)
        if n == "moveTo":
            cmds.append(["M", *pt(c)])
        elif n == "lnTo":
            cmds.append(["L", *pt(c)])
        elif n == "arcTo":
            cmds.append(["A", c.get("wR"), c.get("hR"), c.get("stAng"), c.get("swAng")])
        elif n == "quadBezTo":
            cmds.append(["Q", *[v for q in c for v in (q.get("x"), q.get("y"))]])
        elif n == "cubicBezTo":
            cmds.append(["C", *[v for q in c for v in (q.get("x"), q.get("y"))]])
        elif n == "close":
            cmds.append(["Z"])
    out["c"] = cmds
    return out


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    root = ET.parse(sys.argv[1]).getroot()
    shapes = {}
    for shape in root:
        ah = kid(shape, "ahLst")
        if ah is None or not len(ah):
            continue
        shapes[local(shape.tag)] = {
            "av": guides(kid(shape, "avLst")),
            "gd": guides(kid(shape, "gdLst")),
            "ah": [handle(h) for h in ah],
            "path": [path(p) for p in kid(shape, "pathLst")],
        }
    text = json.dumps(shapes, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    with zipfile.ZipFile(OUT, "w") as z:
        info = zipfile.ZipInfo("handles.json", date_time=(1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, text)
    print(f"{OUT}: {len(shapes)} presets")


if __name__ == "__main__":
    main()
