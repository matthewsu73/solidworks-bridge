#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""SolidWorks COM bridge for agent use (late-bound IDispatch, SW 2025 verified).

All command line lengths are millimetres; SolidWorks API works in metres so the
bridge converts. Every command prints one JSON object on stdout.

    python swbridge.py status
    python swbridge.py new --type part
    python swbridge.py sketch --plane front --shapes "rect 0 0 60 40"
    python swbridge.py extrude --depth 10
    python swbridge.py sketch --plane front --shapes "circle 0 0 8"
    python swbridge.py cut --dir
    python swbridge.py fillet --radius 2 --edges vertical
    python swbridge.py export --out out/plate.step
"""

import argparse
import json
import math
import os
import sys
import traceback

import pythoncom
import win32com.client as win32

# Windows consoles default to a legacy code page; force UTF-8 so the JSON
# payload (feature names are localised, e.g. Chinese) stays valid UTF-8.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

MM = 0.001

DOC_PART, DOC_ASSEMBLY, DOC_DRAWING = 1, 2, 3
DOC_NAME = {1: "part", 2: "assembly", 3: "drawing"}

END_BLIND, END_THROUGH_ALL, END_MID_PLANE = 0, 1, 6
SW_SOLID_BODY = 0  # swBodyType_e.swAllBodies
SW_START_SKETCH_PLANE = 0
FILLET_UNIFORM = 2  # empirically required by FeatureFillet3 in SW2025

TEMPLATE_DIRS = [
    r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2025\templates",
    r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2024\templates",
    r"D:\SolidWorks 2025\SOLIDWORKS\data\Templates",
]
TEMPLATE_NAMES = {
    DOC_PART: ["gb_part.prtdot", "part.prtdot", "Part.prtdot"],
    DOC_ASSEMBLY: ["gb_assembly.asmdot", "assembly.asmdot"],
    DOC_DRAWING: ["gb_a3.drwdot", "gb.drwdot", "A3.drwdot"],
}


# --------------------------------------------------------------------------- #
def val(obj, name, *args):
    """Read a SolidWorks member.

    SolidWorks exposes many zero-argument methods through IDispatch as property
    gets, so calling them raises. With arguments we always call; without, we
    return whatever the attribute already is.
    """
    member = getattr(obj, name)
    if args:
        return member(*args)
    if callable(member) and not hasattr(member, "_oleobj_"):
        return member()
    return member


class Bridge:
    def __init__(self):
        self.sw = win32.Dispatch("SldWorks.Application")
        try:
            self.sw.Visible = True
        except Exception:
            pass

    # -- documents ---------------------------------------------------------- #
    @property
    def model(self):
        m = self.sw.ActiveDoc
        if m is None:
            raise RuntimeError("no active document - run 'new' or 'open' first")
        return m

    def template_for(self, doc_type):
        code = {DOC_PART: 8, DOC_ASSEMBLY: 9, DOC_DRAWING: 10}[doc_type]
        try:
            tpl = val(self.sw, "GetUserPreferenceStringValue", code)
        except Exception:
            tpl = ""
        if tpl and not tpl.startswith("~") and os.path.isfile(tpl):
            return tpl
        for base in TEMPLATE_DIRS:
            if not os.path.isdir(base):
                continue
            for name in TEMPLATE_NAMES[doc_type]:
                candidate = os.path.join(base, name)
                if os.path.isfile(candidate):
                    return candidate
        want = {DOC_PART: ".prtdot", DOC_ASSEMBLY: ".asmdot", DOC_DRAWING: ".drwdot"}[doc_type]
        for base in TEMPLATE_DIRS:
            if not os.path.isdir(base):
                continue
            hits = sorted(f for f in os.listdir(base) if f.lower().endswith(want))
            if hits:
                return os.path.join(base, hits[0])
        return ""

    def new_document(self, doc_type, template=None, paper=None):
        tpl = template or self.template_for(doc_type)
        if doc_type == DOC_DRAWING:
            paper = paper or (6, 0.42, 0.297)  # A3 landscape
            self.sw.NewDocument(tpl, paper[0], paper[1], paper[2])
        else:
            self.sw.NewDocument(tpl, 0, 0, 0)
        m = self.sw.ActiveDoc
        if m is None:
            raise RuntimeError("SolidWorks refused to create a document (template=%r)" % tpl)
        return m

    # -- inspection --------------------------------------------------------- #
    def features(self, m=None):
        m = m or self.model
        out = []
        f = val(m, "FirstFeature")
        while f is not None:
            out.append(f)
            f = val(f, "GetNextFeature")
        return out

    def planes(self, m=None):
        return [f for f in self.features(m) if val(f, "GetTypeName2") == "RefPlane"]

    def last_sketch(self, m=None):
        last = None
        for f in self.features(m):
            if val(f, "GetTypeName2") == "ProfileFeature":
                last = f
        return last

    def doc_info(self, m=None):
        m = m or self.model
        return {
            "title": val(m, "GetTitle"),
            "path": val(m, "GetPathName"),
            "type": DOC_NAME.get(int(val(m, "GetType")), "unknown"),
            "features": [val(f, "Name") for f in self.features(m)],
        }

    def bodies(self, m=None):
        m = m or self.model
        b = m.GetBodies2(SW_SOLID_BODY, False)
        if b is None:
            return []
        if not isinstance(b, (tuple, list)):
            b = [b]
        return list(b)

    def volume_mm3(self, m=None):
        bodies = self.bodies(m)
        if not bodies:
            return None
        mp = bodies[0].GetMassProperties(1000.0)
        return round(float(mp[3]) * 1e9, 2)

    def exact_bbox(self, m=None):
        """Bounding box measured from tessellated faces, in mm.

        GetPartBox / GetBodyBox are approximate for bodies with curved faces
        (a 120x25x120 part can report 124.5x27.9x124.5). Tessellation gives the
        real extent; verified to <0.01 mm against analytic values.
        """
        m = m or self.model
        mn = [float("inf")] * 3
        mx = [float("-inf")] * 3
        faces = 0
        for body in self.bodies(m):
            for face in body.GetFaces():
                faces += 1
                data = face.GetTessTriangles(True)
                for i in range(0, len(data) - 2, 3):
                    for k in range(3):
                        v = data[i + k] * 1000.0
                        if v < mn[k]:
                            mn[k] = v
                        if v > mx[k]:
                            mx[k] = v
        if not faces:
            return None
        return {
            "min": [round(v, 3) for v in mn],
            "max": [round(v, 3) for v in mx],
            "size": [round(mx[k] - mn[k], 3) for k in range(3)],
        }

    # -- sketching ---------------------------------------------------------- #
    def start_sketch(self, plane):
        m = self.model
        planes = self.planes(m)
        key = str(plane).lower()
        index = {"front": 0, "top": 1, "right": 2}.get(key)
        target = None
        if index is not None and len(planes) > index:
            target = planes[index]
        else:
            for f in self.features(m):
                name = val(f, "Name")
                if name and key in name.lower():
                    target = f
                    break
        if target is None:
            raise RuntimeError("plane not found: %s" % plane)
        m.ClearSelection2(True)
        if not target.Select2(False, 0):
            raise RuntimeError("cannot select plane %s" % val(target, "Name"))
        m.SketchManager.InsertSketch(True)
        return val(target, "Name")

    def sketch_shapes(self, shapes):
        m = self.model
        skm = m.SketchManager
        skm.AddToDB = True
        added = []
        try:
            for raw in shapes:
                parts = raw.replace(",", " ").split()
                if not parts:
                    continue
                kind = parts[0].lower()
                num = [float(x) for x in parts[1:]]
                if kind == "circle":
                    cx, cy, r = num[0], num[1], num[2]
                    skm.CreateCircleByRadius(cx * MM, cy * MM, 0, r * MM)
                elif kind == "rect":
                    cx, cy, w, h = num[0], num[1], num[2], num[3]
                    skm.CreateCornerRectangle(
                        (cx - w / 2) * MM, (cy - h / 2) * MM, 0,
                        (cx + w / 2) * MM, (cy + h / 2) * MM, 0)
                elif kind == "line":
                    skm.CreateLine(num[0] * MM, num[1] * MM, 0,
                                   num[2] * MM, num[3] * MM, 0)
                elif kind == "polygon":
                    cx, cy, r, n = num[0], num[1], num[2], int(num[3])
                    rot = math.radians(num[4] if len(num) > 4 else 0)
                    pts = []
                    for i in range(n):
                        a = rot + 2 * math.pi * i / n
                        pts.append(((cx + r * math.cos(a)) * MM,
                                    (cy + r * math.sin(a)) * MM))
                    for i in range(n):
                        x1, y1 = pts[i]
                        x2, y2 = pts[(i + 1) % n]
                        skm.CreateLine(x1, y1, 0, x2, y2, 0)
                elif kind == "slot":
                    cx, cy, length, width = num[0], num[1], num[2], num[3]
                    ang = math.radians(num[4] if len(num) > 4 else 0)
                    r = width / 2.0
                    half = length / 2.0
                    dx, dy = math.cos(ang), math.sin(ang)
                    ax, ay = (cx - dx * half) * MM, (cy - dy * half) * MM
                    bx, by = (cx + dx * half) * MM, (cy + dy * half) * MM
                    skm.CreateLine(ax, ay, 0, bx, by, 0)
                    skm.CreateLine(ax - dy * r, ay + dx * r, 0, bx - dy * r, by + dx * r, 0)
                    skm.CreateLine(ax + dy * r, ay - dx * r, 0, bx + dy * r, by - dx * r, 0)
                else:
                    raise RuntimeError("unknown shape: %s" % kind)
                added.append(kind)
        finally:
            skm.AddToDB = False
        return added

    def end_sketch(self):
        m = self.model
        m.SketchManager.InsertSketch(True)
        sk = self.last_sketch(m)
        return val(sk, "Name") if sk is not None else None

    def sketch_feature(self, name=None):
        m = self.model
        if name:
            for f in self.features(m):
                if val(f, "Name").lower() == name.lower():
                    return f
            for f in self.features(m):
                if name.lower() in val(f, "Name").lower():
                    return f
            raise RuntimeError("sketch not found: %s" % name)
        sk = self.last_sketch(m)
        if sk is None:
            raise RuntimeError("no sketch in document")
        return sk

    # -- features ----------------------------------------------------------- #
    def extrude(self, depth, through_all=False, mid=False, draft=0.0,
                flip=False, sketch=None, merge=True):
        m = self.model
        sk = self.sketch_feature(sketch)
        m.ClearSelection2(True)
        if not sk.Select2(False, 0):
            raise RuntimeError("cannot select sketch %s" % val(sk, "Name"))
        T1 = END_THROUGH_ALL if through_all else (END_MID_PLANE if mid else END_BLIND)
        D1 = 0.0 if through_all else depth * MM
        feat = m.FeatureManager.FeatureExtrusion3(
            True, flip, False,
            T1, 0, D1, 0.0,
            False, False, False, False,
            draft, 0.0,
            False, False, False, False,
            merge, False, False,
            SW_START_SKETCH_PLANE, 0.0, False)
        val(m, "EditRebuild3")
        m.ClearSelection2(True)
        if feat is None:
            raise RuntimeError("extrusion failed")
        return {"ok": True, "feature": val(feat, "Name")}

    def cut(self, depth=10.0, through_all=True, dir_=False, flip=False,
            draft=0.0, sketch=None):
        m = self.model
        sk = self.sketch_feature(sketch)
        m.ClearSelection2(True)
        if not sk.Select2(False, 0):
            raise RuntimeError("cannot select sketch %s" % val(sk, "Name"))
        T1 = END_THROUGH_ALL if through_all else END_BLIND
        D1 = 0.0 if through_all else depth * MM
        feat = m.FeatureManager.FeatureCut4(
            True, flip, dir_,
            T1, 0, D1, 0.0,
            False, False, False, False,
            draft, 0.0,
            False, False, False, False,
            False, False, False,
            False, False, False,
            SW_START_SKETCH_PLANE, 0.0, False,
            False)
        val(m, "EditRebuild3")
        m.ClearSelection2(True)
        if feat is None:
            raise RuntimeError("cut failed - point it at the body with --dir or --flip")
        return {"ok": True, "feature": val(feat, "Name")}

    def revolve(self, angle=360.0, sketch=None, flip=False):
        m = self.model
        sk = self.sketch_feature(sketch)
        m.ClearSelection2(True)
        sk.Select2(False, 0)
        feat = m.FeatureManager.FeatureRevolve2(
            True, True, False, False, flip, False,
            0, 0, math.radians(angle), 0.0,
            False, False, 0.0, 0.0,
            0, 0.0, 0.0,
            True, False, False)
        val(m, "EditRebuild3")
        m.ClearSelection2(True)
        if feat is None:
            raise RuntimeError("revolve failed - the sketch needs a centreline/axis")
        return {"ok": True, "feature": val(feat, "Name")}

    # -- edge selection ----------------------------------------------------- #
    def edges(self, m=None):
        out = []
        for body in self.bodies(m):
            got = body.GetEdges()
            if got:
                out.extend(got)
        return out

    def edge_points(self, edge):
        v1 = val(edge, "GetStartVertex")
        v2 = val(edge, "GetEndVertex")
        p1 = list(val(v1, "GetPoint")) if v1 is not None else None
        p2 = list(val(v2, "GetPoint")) if v2 is not None else None
        return p1, p2

    def select_edges(self, spec="all"):
        """spec: all | vertical | horizontal | circular | near:x,y,z;x,y,z"""
        m = self.model
        spec = (spec or "all").strip().lower()
        m.ClearSelection2(True)
        count = 0
        points = None
        if spec.startswith("near:"):
            points = []
            for chunk in spec[5:].split(";"):
                chunk = chunk.strip()
                if not chunk:
                    continue
                c = [float(x) for x in chunk.split(",")]
                points.append([v * MM for v in c])
        for edge in self.edges(m):
            p1, p2 = self.edge_points(edge)
            keep = True
            if points is not None:
                keep = False
                for p in points:
                    for endpoint in (p1, p2):
                        if endpoint and all(abs(endpoint[i] - p[i]) < 0.5 * MM for i in range(3)):
                            keep = True
                    if not keep and p1 and p2:
                        mid = [(p1[i] + p2[i]) / 2 for i in range(3)]
                        if all(abs(mid[i] - p[i]) < 0.5 * MM for i in range(3)):
                            keep = True
            elif spec == "circular":
                keep = p1 is None or p2 is None or (
                    abs(p1[0] - p2[0]) < 1e-9 and abs(p1[1] - p2[1]) < 1e-9
                    and abs(p1[2] - p2[2]) < 1e-9)
            elif spec in ("vertical", "horizontal"):
                if p1 is None or p2 is None:
                    keep = False
                else:
                    dz = abs(p2[2] - p1[2])
                    length = math.sqrt(sum((p2[i] - p1[i]) ** 2 for i in range(3)))
                    ratio = dz / length if length > 1e-9 else 0
                    keep = ratio > 0.98 if spec == "vertical" else ratio < 0.02
            if keep and edge.Select2(True, 0):
                count += 1
        return count

    def fillet(self, radius, edges="vertical"):
        m = self.model
        n = self.select_edges(edges)
        if n == 0:
            raise RuntimeError("no edges matched %r" % edges)
        feat = m.FeatureManager.FeatureFillet3(
            FILLET_UNIFORM, radius * MM, 0.0, 0.0, 0, 0, 0,
            None, None, None, None, None, None, None)
        val(m, "EditRebuild3")
        m.ClearSelection2(True)
        if feat is None:
            raise RuntimeError("fillet failed (%d edges selected)" % n)
        return {"ok": True, "feature": val(feat, "Name"), "edges": n}

    def chamfer(self, distance, angle=45.0, edges="vertical"):
        m = self.model
        n = self.select_edges(edges)
        if n == 0:
            raise RuntimeError("no edges matched %r" % edges)
        feat = m.FeatureManager.InsertFeatureChamfer(
            0, 0, distance * MM, math.radians(angle), 0.0, 0.0, 0.0, 0.0)
        val(m, "EditRebuild3")
        m.ClearSelection2(True)
        if feat is None:
            raise RuntimeError("chamfer failed (%d edges selected)" % n)
        return {"ok": True, "feature": val(feat, "Name"), "edges": n}

    # -- io ----------------------------------------------------------------- #
    def export(self, out_path, fmt=None):
        m = self.model
        out_path = os.path.abspath(out_path)
        folder = os.path.dirname(out_path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        code = m.SaveAs3(out_path, 0, 1)
        if not os.path.isfile(out_path) or os.path.getsize(out_path) == 0:
            raise RuntimeError("export failed for %s (code=%r)" % (out_path, code))
        return {
            "ok": True,
            "path": out_path,
            "format": (fmt or os.path.splitext(out_path)[1].lstrip(".")).lower(),
            "bytes": os.path.getsize(out_path),
        }

    def info(self):
        m = self.model
        out = self.doc_info(m)
        out["ok"] = True
        try:
            box = list(m.GetPartBox(True))
            out["bbox_mm"] = {
                "min": [round(v / MM, 3) for v in box[:3]],
                "max": [round(v / MM, 3) for v in box[3:]],
                "size": [round((box[i + 3] - box[i]) / MM, 3) for i in range(3)],
            }
        except Exception:
            pass
        try:
            exact = self.exact_bbox(m)
            if exact is not None:
                out["bbox_mm_exact"] = exact
        except Exception:
            pass
        bodies = self.bodies(m)
        out["bodies"] = len(bodies)
        if bodies:
            mp = bodies[0].GetMassProperties(1000.0)
            out["mass_properties"] = {
                "volume_mm3": round(float(mp[3]) * 1e9, 2),
                "area_mm2": round(float(mp[4]) * 1e6, 2),
                "mass_kg_at_1000kgm3": round(float(mp[5]), 4),
                "center_mm": [round(float(v) / MM, 3) for v in mp[0:3]],
            }
        return out

    def screenshot(self, out_path, width=1400, height=900):
        m = self.model
        out_path = os.path.abspath(out_path)
        folder = os.path.dirname(out_path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        val(m, "ViewZoomtofit2")
        if not m.SaveBMP(out_path, width, height):
            raise RuntimeError("SaveBMP failed")
        return {"ok": True, "path": out_path}


# --------------------------------------------------------------------------- #
def parse_shapes(args):
    if args.shapes_file:
        if args.shapes_file == "-":
            raw = sys.stdin.read()
        else:
            with open(args.shapes_file, "r", encoding="utf-8-sig") as fh:
                raw = fh.read()
    else:
        raw = args.shapes or ""
    chunks = []
    for line in raw.replace(";", "\n").splitlines():
        line = line.strip()
        if line:
            chunks.append(line)
    return chunks


def cmd_status(b, a):
    out = {"ok": True, "revision": val(b.sw, "RevisionNumber"), "docs": b.sw.GetDocumentCount}
    if b.sw.ActiveDoc is not None:
        out["active"] = b.doc_info()
    return out


def cmd_new(b, a):
    t = {"part": DOC_PART, "assembly": DOC_ASSEMBLY, "drawing": DOC_DRAWING}[a.type]
    paper = None
    if a.paper_size is not None:
        paper = (a.paper_size, a.paper_w, a.paper_h)
    b.new_document(t, a.template, paper)
    return {"ok": True, "active": b.doc_info()}


def cmd_open(b, a):
    path = os.path.abspath(a.path)
    if not os.path.isfile(path):
        raise RuntimeError("file not found: %s" % path)
    dtype = {".sldprt": 1, ".sldasm": 2, ".slddrw": 3}.get(
        os.path.splitext(path)[1].lower(), 1)
    errors = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    doc = b.sw.OpenDoc6(path, dtype, 1, "", errors, warnings)
    if doc is None:
        raise RuntimeError("OpenDoc6 failed (errors=%s warnings=%s)"
                           % (errors.value, warnings.value))
    return {"ok": True, "active": b.doc_info(doc)}


def cmd_save(b, a):
    m = b.model
    if a.path:
        path = os.path.abspath(a.path)
        folder = os.path.dirname(path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        m.SaveAs3(path, 0, 1)
        if not os.path.isfile(path):
            raise RuntimeError("save failed")
        return {"ok": True, "path": path}
    errors, warnings = 0, 0
    m.Save3(1, errors, warnings)
    return {"ok": True, "path": val(m, "GetPathName")}


def cmd_close(b, a):
    m = b.sw.ActiveDoc
    if m is None:
        return {"ok": True, "closed": ""}
    title = val(m, "GetTitle")
    b.sw.CloseDoc(title)
    return {"ok": True, "closed": title}


def cmd_sketch(b, a):
    shapes = parse_shapes(a)
    plane_name = b.start_sketch(a.plane)
    added = b.sketch_shapes(shapes)
    name = b.end_sketch()
    return {"ok": True, "sketch": name, "plane": plane_name, "shapes": added}


def cmd_extrude(b, a):
    return b.extrude(a.depth, through_all=a.through_all, mid=a.mid,
                     draft=a.draft, flip=a.flip, sketch=a.sketch,
                     merge=not a.no_merge)


def cmd_cut(b, a):
    return b.cut(a.depth, through_all=not a.blind, dir_=a.dir, flip=a.flip,
                 draft=a.draft, sketch=a.sketch)


def cmd_revolve(b, a):
    return b.revolve(a.angle, sketch=a.sketch, flip=a.flip)


def cmd_fillet(b, a):
    return b.fillet(a.radius, edges=a.edges)


def cmd_chamfer(b, a):
    return b.chamfer(a.distance, a.angle, edges=a.edges)


def cmd_export(b, a):
    return b.export(a.out, a.format)


def cmd_info(b, a):
    return b.info()


def cmd_rebuild(b, a):
    m = b.model
    val(m, "EditRebuild3")
    val(m, "ViewZoomtofit2")
    return {"ok": True, "volume_mm3": b.volume_mm3()}


def cmd_screenshot(b, a):
    return b.screenshot(a.out, a.width, a.height)


def cmd_macro(b, a):
    path = os.path.abspath(a.path)
    if not os.path.isfile(path):
        raise RuntimeError("macro not found: %s" % path)
    errors = 0
    ok = b.sw.RunMacro2(path, a.module, a.method, 0, errors)
    if not ok:
        raise RuntimeError("macro failed (errors=%s)" % errors)
    return {"ok": True, "macro": path}


def cmd_exec(b, a):
    code = a.code if a.code else sys.stdin.read()
    import io
    from contextlib import redirect_stdout

    g = {"sw": b.sw, "bridge": b, "model": b.model, "MM": MM,
         "json": json, "math": math, "val": val}
    buf = io.StringIO()
    with redirect_stdout(buf):
        exec(compile(code, "<exec>", "exec"), g)
    out = {"ok": True, "stdout": buf.getvalue()}
    if "result" in g:
        out["result"] = g["result"]
    return out


def build_parser():
    p = argparse.ArgumentParser(description="SolidWorks COM bridge (mm units)")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("status")

    n = sub.add_parser("new")
    n.add_argument("--type", default="part", choices=["part", "assembly", "drawing"])
    n.add_argument("--template", default=None)
    n.add_argument("--paper-size", type=int, default=6)
    n.add_argument("--paper-w", type=float, default=0.42)
    n.add_argument("--paper-h", type=float, default=0.297)

    o = sub.add_parser("open")
    o.add_argument("path")

    s = sub.add_parser("save")
    s.add_argument("--path", default=None)

    sub.add_parser("close")

    sk = sub.add_parser("sketch")
    sk.add_argument("--plane", default="front")
    sk.add_argument("--shapes", default="")
    sk.add_argument("--shapes-file", default=None)

    e = sub.add_parser("extrude")
    e.add_argument("--depth", type=float, default=10.0)
    e.add_argument("--through-all", action="store_true")
    e.add_argument("--mid", action="store_true")
    e.add_argument("--draft", type=float, default=0.0)
    e.add_argument("--flip", action="store_true")
    e.add_argument("--sketch", default=None)
    e.add_argument("--no-merge", action="store_true")

    c = sub.add_parser("cut")
    c.add_argument("--depth", type=float, default=10.0)
    c.add_argument("--blind", action="store_true")
    c.add_argument("--dir", action="store_true")
    c.add_argument("--flip", action="store_true")
    c.add_argument("--draft", type=float, default=0.0)
    c.add_argument("--sketch", default=None)

    r = sub.add_parser("revolve")
    r.add_argument("--angle", type=float, default=360.0)
    r.add_argument("--sketch", default=None)
    r.add_argument("--flip", action="store_true")

    f = sub.add_parser("fillet")
    f.add_argument("--radius", type=float, required=True)
    f.add_argument("--edges", default="vertical")

    ch = sub.add_parser("chamfer")
    ch.add_argument("--distance", type=float, required=True)
    ch.add_argument("--angle", type=float, default=45.0)
    ch.add_argument("--edges", default="vertical")

    ex = sub.add_parser("export")
    ex.add_argument("--out", required=True)
    ex.add_argument("--format", default=None)

    sub.add_parser("info")
    sub.add_parser("rebuild")

    sc = sub.add_parser("screenshot")
    sc.add_argument("--out", required=True)
    sc.add_argument("--width", type=int, default=1400)
    sc.add_argument("--height", type=int, default=900)

    mc = sub.add_parser("macro")
    mc.add_argument("path")
    mc.add_argument("--module", default="main")
    mc.add_argument("--method", default="main")

    ec = sub.add_parser("exec")
    ec.add_argument("--code", default=None)

    return p


COMMANDS = {
    "status": cmd_status, "new": cmd_new, "open": cmd_open, "save": cmd_save,
    "close": cmd_close, "sketch": cmd_sketch, "extrude": cmd_extrude,
    "cut": cmd_cut, "revolve": cmd_revolve, "fillet": cmd_fillet,
    "chamfer": cmd_chamfer, "export": cmd_export, "info": cmd_info,
    "rebuild": cmd_rebuild, "screenshot": cmd_screenshot, "macro": cmd_macro,
    "exec": cmd_exec,
}


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        bridge = Bridge()
        result = COMMANDS[args.cmd](bridge, args)
    except Exception as exc:  # noqa: BLE001 - reported to the caller as JSON
        print(json.dumps({
            "ok": False,
            "error": "%s: %s" % (type(exc).__name__, exc),
            "traceback": traceback.format_exc(limit=2),
        }, ensure_ascii=False, indent=2))
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
