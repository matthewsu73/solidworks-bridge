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
SW_SOLID_BODY = 0    # GetBodies2(0) -> solid bodies (verified SW2025)
SW_SURFACE_BODY = 1  # GetBodies2(1) -> surface / sheet bodies (verified SW2025)
BODY_KINDS = {"solid": SW_SOLID_BODY, "surface": SW_SURFACE_BODY}
# swBodyOperationType_e, verified on SW2025 via FeatureManager.InsertCombineFeature
COMBINE_OPS = {"add": 15901, "subtract": 15902, "intersect": 15903}
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

    def bodies(self, m=None, kind="all"):
        """Bodies in the active part.

        kind: all | solid | surface. SolidWorks stores solid and sheet bodies in
        separate collections, so both have to be queried and merged.
        """
        m = m or self.model
        if kind == "all":
            types = [SW_SOLID_BODY, SW_SURFACE_BODY]
        else:
            try:
                types = [BODY_KINDS[kind]]
            except KeyError:
                raise RuntimeError("unknown body kind: %s (all|solid|surface)" % kind)
        out = []
        for t in types:
            raw = m.GetBodies2(t, False)
            if raw is None:
                continue
            if not isinstance(raw, (tuple, list)):
                raw = [raw]
            out.extend(raw)
        return out

    def body_kind(self, body):
        try:
            return "surface" if int(val(body, "GetType")) == SW_SURFACE_BODY else "solid"
        except Exception:
            return "unknown"

    def solid_bodies(self, m=None):
        return self.bodies(m, "solid")

    def surface_bodies(self, m=None):
        return self.bodies(m, "surface")

    def volume_mm3(self, m=None):
        solids = self.solid_bodies(m)
        if not solids:
            return None
        total = 0.0
        for body in solids:
            total += float(body.GetMassProperties(1000.0)[3])
        return round(total * 1e9, 2)

    def surface_area_mm2(self, m=None, kind="surface"):
        """Area from IFace2.GetArea (m^2) summed over faces, in mm^2.

        GetMassProperties() is not meaningful for open sheet bodies; this is.
        """
        total = 0.0
        for body in self.bodies(m, kind):
            for face in body.GetFaces():
                total += float(face.GetArea)
        return round(total * 1e6, 2)

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

    def body_box_mm(self, body):
        """Exact [xmin,ymin,zmin,xmax,ymax,zmax] for one body, in mm."""
        mn = [float("inf")] * 3
        mx = [float("-inf")] * 3
        seen = False
        for face in body.GetFaces():
            data = face.GetTessTriangles(True)
            if not data:
                continue
            seen = True
            for i in range(0, len(data) - 2, 3):
                for k in range(3):
                    v = data[i + k] * 1000.0
                    if v < mn[k]:
                        mn[k] = v
                    if v > mx[k]:
                        mx[k] = v
        if not seen:
            return None
        return mn + mx

    def body_volume_mm3(self, body):
        try:
            return round(float(body.GetMassProperties(1000.0)[3]) * 1e9, 2)
        except Exception:
            return 0.0

    # -- sketching ---------------------------------------------------------- #
    def find_plane(self, plane):
        """front|top|right by position, otherwise fuzzy name match."""
        m = self.model
        planes = self.planes(m)
        key = str(plane).lower()
        index = {"front": 0, "top": 1, "right": 2}.get(key)
        if index is not None and len(planes) > index:
            return planes[index]
        for f in self.features(m):
            name = val(f, "Name")
            if name and key in name.lower():
                return f
        raise RuntimeError("plane not found: %s" % plane)

    def plane_offset(self, base="front", distance=10.0):
        m = self.model
        target = self.find_plane(base)
        before = self._feature_names(m)
        m.ClearSelection2(True)
        if not target.Select2(False, 0):
            raise RuntimeError("cannot select plane %s" % val(target, "Name"))
        flip = distance < 0
        m.CreatePlaneAtOffset3(abs(distance) * MM, flip, True)
        name = self._created_feature(before, m, "offset plane")
        m.ClearSelection2(True)
        return {"ok": True, "plane": name, "base": val(target, "Name")}

    def start_sketch(self, plane):
        m = self.model
        target = self.find_plane(plane)
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
                elif kind == "centerline":
                    skm.CreateCenterLine(num[0] * MM, num[1] * MM, 0,
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

    # -- surfaces (all verified on SOLIDWORKS 2025) ------------------------- #
    def _feature_names(self, m=None):
        return [val(f, "Name") for f in self.features(m)]

    def _created_feature(self, before, m=None, what="feature", hint=""):
        """Rebuild and return the newest feature, failing if none appeared.

        InsertSweepSurface3 / InsertOffsetSurface / InsertExtendSurface /
        InsertSewRefSurface return None on success, so the feature tree is the
        only reliable success signal.
        """
        m = m or self.model
        val(m, "EditRebuild3")
        names = self._feature_names(m)
        if len(names) <= len(before):
            msg = "%s failed - no feature was created" % what
            if hint:
                msg += " (" + hint + ")"
            raise RuntimeError(msg)
        return names[-1]

    def surface_plane(self, sketch=None):
        m = self.model
        sk = self.sketch_feature(sketch)
        before = self._feature_names(m)
        m.ClearSelection2(True)
        if not sk.Select2(False, 0):
            raise RuntimeError("cannot select sketch %s" % val(sk, "Name"))
        val(m, "InsertPlanarRefSurface")
        name = self._created_feature(before, m, "planar surface")
        m.ClearSelection2(True)
        return {"ok": True, "feature": name}

    def surface_extrude(self, depth, flip=False, sketch=None):
        m = self.model
        sk = self.sketch_feature(sketch)
        before = self._feature_names(m)
        m.ClearSelection2(True)
        if not sk.Select2(False, 0):
            raise RuntimeError("cannot select sketch %s" % val(sk, "Name"))
        m.FeatureManager.FeatureExtruRefSurface3(
            True, flip, 0, 0.0, 0, 0, depth * MM, 0.0,
            False, False, False, False, 0.0, 0.0,
            False, False, False, False, False, False, True, False)
        name = self._created_feature(before, m, "surface extrude")
        m.ClearSelection2(True)
        return {"ok": True, "feature": name}

    def surface_revolve(self, angle=360.0, sketch=None, flip=False):
        m = self.model
        sk = self.sketch_feature(sketch)
        before = self._feature_names(m)
        m.ClearSelection2(True)
        if not sk.Select2(False, 0):
            raise RuntimeError("cannot select sketch %s" % val(sk, "Name"))
        m.FeatureManager.FeatureRevolve2(
            True, False, False, False, flip, False,
            0, 0, math.radians(angle), 0.0,
            False, False, 0.0, 0.0,
            0, 0.0, 0.0,
            True, False, False)
        name = self._created_feature(
            before, m, "surface revolve",
            "check the sketch has a centreline and the profile does not sit on it")
        m.ClearSelection2(True)
        return {"ok": True, "feature": name}

    def surface_loft(self, sketches):
        m = self.model
        if len(sketches) < 2:
            raise RuntimeError("surface-loft needs at least two sketches")
        feats = [self.sketch_feature(name) for name in sketches]
        before = self._feature_names(m)
        m.ClearSelection2(True)
        for f in feats:
            if not f.Select2(True, 1):
                raise RuntimeError("cannot select sketch %s" % val(f, "Name"))
        m.InsertLoftRefSurface2(False, False, False, 0.0, 0, 0)
        name = self._created_feature(before, m, "surface loft")
        m.ClearSelection2(True)
        return {"ok": True, "feature": name, "sketches": len(feats)}

    def surface_sweep(self, outline, path):
        m = self.model
        outline_sk = self.sketch_feature(outline)
        path_sk = self.sketch_feature(path)
        before = self._feature_names(m)
        m.ClearSelection2(True)
        # Outline takes mark 1, path takes mark 4 (verified on SW2025; the
        # reverse order creates a feature with the wrong geometry).
        if not outline_sk.Select2(True, 1):
            raise RuntimeError("cannot select outline sketch %s" % val(outline_sk, "Name"))
        if not path_sk.Select2(True, 4):
            raise RuntimeError("cannot select path sketch %s" % val(path_sk, "Name"))
        m.FeatureManager.InsertSweepSurface3(
            False, 0, False, True, 0, 0, 0,
            True, True, 0.0, False, False, 0.0, False)
        name = self._created_feature(
            before, m, "surface sweep",
            "the path must cross the profile plane, e.g. a line on the top plane going along Z")
        m.ClearSelection2(True)
        return {"ok": True, "feature": name}

    def thicken(self, thickness, faces="all", merge=True):
        """Thicken surface faces into a solid.

        merge=False keeps the result as its own body, which is what a mould
        cavity needs (the shell must stay separate so combine can subtract it).
        """
        m = self.model
        before = self._feature_names(m)
        n = self.select_faces(faces, kind="surface", mark=1)
        if n == 0:
            raise RuntimeError("no surface faces matched %r" % faces)
        m.FeatureManager.FeatureBossThicken(
            thickness * MM, 0, 0, False, bool(merge), False, False)
        name = self._created_feature(before, m, "thicken")
        m.ClearSelection2(True)
        return {
            "ok": True, "feature": name, "faces": n,
            "volume_mm3": self.volume_mm3(m),
            "solid_bodies": len(self.solid_bodies(m)),
        }

    def offset_surface(self, distance, faces="all"):
        m = self.model
        before = self._feature_names(m)
        n = self.select_faces(faces, kind="surface", mark=0)
        if n == 0:
            raise RuntimeError("no surface faces matched %r" % faces)
        m.InsertOffsetSurface(distance * MM, False)
        name = self._created_feature(before, m, "offset surface")
        m.ClearSelection2(True)
        return {"ok": True, "feature": name, "faces": n}

    def sew(self, bodies="surface", tolerance=0.01):
        m = self.model
        targets = self.bodies(m, bodies)
        if not targets:
            raise RuntimeError("no %s bodies to sew" % bodies)
        before = self._feature_names(m)
        before_bodies = len(self.bodies(m))
        m.ClearSelection2(True)
        for body in targets:
            if not body.Select(True, 1):
                raise RuntimeError("cannot select body %s" % val(body, "Name"))
        m.FeatureManager.InsertSewRefSurface(
            False, False, True, tolerance * MM, 0.0)
        name = self._created_feature(before, m, "sew")
        m.ClearSelection2(True)
        return {
            "ok": True, "feature": name, "sewn": len(targets),
            "bodies_before": before_bodies,
            "bodies_after": len(self.bodies(m)),
            "surface_bodies": len(self.surface_bodies(m)),
        }

    def extend_surface(self, distance, edges="all"):
        m = self.model
        before = self._feature_names(m)
        n = self.select_edges(edges, kind="surface")
        if n == 0:
            raise RuntimeError("no surface edges matched %r" % edges)
        m.InsertExtendSurface(True, 0, distance * MM)
        name = self._created_feature(before, m, "extend surface")
        m.ClearSelection2(True)
        return {"ok": True, "feature": name, "edges": n}

    # -- booleans ----------------------------------------------------------- #
    def combine(self, operation="subtract", target="largest", tools="all"):
        """Boolean between solid bodies.

        Subtract keeps `--target` and removes `--tools` from it, which is the
        mould workflow: target = blank block, tools = cavity solid.
        """
        m = self.model
        key = (operation or "subtract").strip().lower()
        if key not in COMBINE_OPS:
            raise RuntimeError("unknown operation %r (add|subtract|intersect)" % operation)
        solids = self.solid_bodies(m)
        if len(solids) < 2:
            raise RuntimeError(
                "combine needs at least two solid bodies, found %d "
                "(extrude with --no-merge / thicken separate surfaces)" % len(solids))

        def center(body):
            box = self.body_box_mm(body)
            if box is None:
                return None
            return [(box[i] + box[i + 3]) / 2.0 for i in range(3)]

        def near(spec, pool):
            point = [float(x) * MM for x in spec[5:].split(",")[:3]]
            best, best_d = None, None
            for body in pool:
                c = center(body)
                if c is None:
                    continue
                d = sum((c[i] - point[i]) ** 2 for i in range(3))
                if best_d is None or d < best_d:
                    best_d, best = d, body
            return best

        def pick(spec, pool, what):
            spec = (spec or "").strip().lower()
            if spec.startswith("#"):
                try:
                    i = int(spec[1:])
                except ValueError:
                    raise RuntimeError("bad %s %r, use #index" % (what, spec))
                if i < 0 or i >= len(solids):
                    raise RuntimeError("%s index out of range: %s (have %d bodies)"
                                       % (what, spec, len(solids)))
                if not any(solids[i] is b for b in pool):
                    raise RuntimeError("%s %s is not a usable %s body"
                                       % (what, spec, what))
                return solids[i]
            if spec.startswith("near:"):
                body = near(spec, pool)
                if body is None:
                    raise RuntimeError("no %s body matched %r" % (what, spec))
                return body
            if spec in ("largest", "biggest", "max"):
                return max(pool, key=self.body_volume_mm3)
            if spec in ("smallest", "min"):
                return min(pool, key=self.body_volume_mm3)
            if spec == "first":
                return pool[0]
            if spec == "last":
                return pool[-1]
            raise RuntimeError("bad %s %r (largest|smallest|first|last|#i|near:x,y,z)"
                               % (what, spec))

        main = pick(target, solids, "target")
        if (tools or "all").strip().lower() == "all":
            picked = [b for b in solids if b is not main]
        else:
            picked = [pick(tools, [b for b in solids if b is not main], "tool")]

        def index_of(body):
            for i, b in enumerate(solids):
                if b is body:
                    return i
            return -1

        def label(body):
            name = val(body, "Name") or ""
            i = index_of(body)
            return name if name else "#%d" % i

        before = self._feature_names(m)
        vol_before = [self.body_volume_mm3(b) for b in solids]
        main_label = label(main)
        tool_labels = [label(b) for b in picked]
        m.ClearSelection2(True)
        for body in [main] + picked:
            if not body.Select(True, 1):
                raise RuntimeError("cannot select body %s" % val(body, "Name"))
        arr = win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, picked)
        m.FeatureManager.InsertCombineFeature(COMBINE_OPS[key], main, arr)
        name = self._created_feature(before, m, "combine (%s)" % key)
        m.ClearSelection2(True)
        return {
            "ok": True, "feature": name, "operation": key,
            "target": main_label, "target_volume_mm3": vol_before[index_of(main)],
            "tools": tool_labels,
            "volume_before_mm3": vol_before,
            "volume_mm3": self.volume_mm3(m),
            "solid_bodies": len(self.solid_bodies(m)),
        }

# -- surface trim / fill / boundary ------------------------------------ #
    def _select_body_by_name(self, body, append, mark, doc_type="SURFACEBODY"):
        """Select a whole body through its name (avoids face-level ambiguity)."""
        m = self.model
        callout = win32.VARIANT(pythoncom.VT_DISPATCH, None)
        return bool(m.Extension.SelectByID2(
            val(body, "Name") or "", doc_type, 0.0, 0.0, 0.0,
            append, mark, callout, 0))

    @staticmethod
    def _body_center_mm(body):
        try:
            box = list(val(body, "GetBodyBox"))
        except Exception:
            return None
        if box is None or len(box) < 6:
            return None
        return [(box[i] + box[i + 3]) / 2.0 for i in range(3)]

    def trim_surface(self, tool="last", target="first", mode="standard",
                     sew=False, tool_mark=1, target_mark=4):
        """Trim `target` surface body with a `tool` surface body (SW UI: trim surface).
        SolidWorks distinguishes trim tool and target by selection marks; defaults
        follow the UI convention (tool=1, target=4). Swap via tool-mark/target-mark."""
        m = self.model
        mode = (mode or "standard").strip().lower()
        if mode not in ("standard", "mutual"):
            raise RuntimeError("bad --mode %r (standard|mutual)" % mode)
        surfaces = self.surface_bodies(m)
        if len(surfaces) < 2:
            raise RuntimeError(
                "trim-surface needs at least two surface bodies, found %d"
                % len(surfaces))

        def pick(spec, default_i, what):
            spec = (spec or "").strip().lower()
            if spec.startswith("#"):
                try:
                    i = int(spec[1:])
                except ValueError:
                    raise RuntimeError("bad %s %r, use #index" % (what, spec))
                if i < 0 or i >= len(surfaces):
                    raise RuntimeError("%s index out of range: %s"
                                       % (what, spec))
                return surfaces[i]
            if spec.startswith("near:"):
                try:
                    point = [float(x) * MM for x in spec[5:].split(",")[:3]]
                except ValueError:
                    raise RuntimeError("bad %s %r, use near:x,y,z" % (what, spec))
                best, best_d = None, None
                for body in surfaces:
                    c = self._body_center_mm(body)
                    if c is None:
                        continue
                    d = sum((c[k] - point[k]) ** 2 for k in range(3))
                    if best_d is None or d < best_d:
                        best_d, best = d, body
                if best is None:
                    raise RuntimeError("no %s body near %r" % (what, spec))
                return best
            if spec in ("first", "last"):
                return surfaces[0 if spec == "first" else -1]
            if spec in ("", "default"):
                return surfaces[default_i]
            for body in surfaces:
                if spec in (val(body, "Name") or "").lower():
                    return body
            raise RuntimeError(
                "no %s body matched %r (first|last|#i|near:x,y,z|name)"
                % (what, spec))

        tool_body = pick(tool, -1, "tool")
        target_body = pick(target, 0, "target")
        if tool_body is target_body:
            raise RuntimeError("tool and target resolved to the same body")

        before = self._feature_names(m)
        m.ClearSelection2(True)
        self._select_body_by_name(tool_body, False, tool_mark)
        self._select_body_by_name(target_body, True, target_mark)
        if not m.PreTrimSurface(mode == "mutual"):
            raise RuntimeError(
                "PreTrimSurface rejected the selection - try swapping "
                "--tool-mark/--target-mark")
        m.PostTrimSurface(bool(sew))
        name = self._created_feature(
            before, m, "trim surface",
            "no trim feature appeared - try swapped marks (--tool-mark %d "
            "--target-mark %d)" % (target_mark, tool_mark))
        m.ClearSelection2(True)
        return {
            "ok": True, "feature": name, "mode": mode,
            "tool": val(tool_body, "Name"), "target": val(target_body, "Name"),
            "surface_bodies": len(self.surface_bodies(m)),
        }

    def fill_surface(self, edges="all", resolution=1):
        """Fill a patch surface on the selected closed edge loop (fill surface)."""
        m = self.model
        n = self.select_edges(edges)
        if n == 0:
            raise RuntimeError("no edges matched %r (all|near:x,y,z)" % edges)
        before = self._feature_names(m)
        m.FeatureManager.InsertFillSurface(int(resolution))
        name = self._created_feature(
            before, m, "fill surface",
            "select a closed loop of edges/sketches")
        m.ClearSelection2(True)
        return {"ok": True, "feature": name, "edges": n}

    def boundary_surface(self, sketches=None, from_feature=None):
        """Boundary surface: SW2025 API has no direct create call for boundary.

        Supported equivalent: build a loft surface then convert it to a net
        blend (boundary) feature via ConvertLoftOrSweepToNetBlend, or convert
        an existing loft/sweep surface feature (--from-feature)."""
        m = self.model
        if from_feature:
            loft_name = from_feature
        else:
            if not sketches or len(sketches) < 2:
                raise RuntimeError(
                    "boundary-surface needs --sketches s1,s2 (or more) "
                    "or --from-feature <existing loft/sweep feature>")
            loft_name = val(self.surface_loft(sketches), "feature")
        m.ClearSelection2(True)
        callout = win32.VARIANT(pythoncom.VT_DISPATCH, None)
        ok = m.Extension.SelectByID2(
            loft_name, "BODYFEATURE", 0.0, 0.0, 0.0, False, 1, callout, 0)
        if not ok:
            raise RuntimeError(
                "cannot select feature %r for conversion (has it been "
                "created already?)" % loft_name)
        before = self._feature_names(m)
        m.FeatureManager.ConvertLoftOrSweepToNetBlend()
        name = self._created_feature(
            before, m, "boundary surface",
            "conversion produced no feature - verify the source is a "
            "loft/sweep surface")
        m.ClearSelection2(True)
        return {"ok": True, "feature": name, "from": loft_name,
                "surface_bodies": len(self.surface_bodies(m))}

    # -- edge / face selection ---------------------------------------------- #
    def edges(self, m=None, kind="all"):
        out = []
        for body in self.bodies(m, kind):
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

    def select_edges(self, spec="all", kind="all"):
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
        for edge in self.edges(m, kind):
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

    def faces(self, m=None, kind="all"):
        out = []
        for body in self.bodies(m, kind):
            got = body.GetFaces()
            if got:
                out.extend(got)
        return out

    def face_center_mm(self, face):
        """Mean of the face's tessellation vertices, in mm."""
        data = face.GetTessTriangles(True)
        if not data:
            return None
        n = 0
        acc = [0.0, 0.0, 0.0]
        for i in range(0, len(data) - 2, 3):
            for k in range(3):
                acc[k] += float(data[i + k]) * 1000.0
            n += 1
        return [v / n for v in acc]

    def select_faces(self, spec="all", kind="surface", mark=1):
        """spec: all | near:x,y,z (mm).

        near: picks the single face whose tessellation centroid is closest to
        the given point - the practical way to pick a face without a UI.
        """
        m = self.model
        spec = (spec or "all").strip().lower()
        m.ClearSelection2(True)
        candidates = self.faces(m, kind)
        chosen = []
        if spec.startswith("near:"):
            try:
                target = [float(x) * MM for x in spec[5:].split(",")[:3]]
            except ValueError:
                raise RuntimeError("bad face point %r, use near:x,y,z" % spec)
            if len(target) != 3:
                raise RuntimeError("bad face point %r, use near:x,y,z" % spec)
            best = None
            best_d = None
            for face in candidates:
                c = self.face_center_mm(face)
                if c is None:
                    continue
                d = sum((c[i] - target[i]) ** 2 for i in range(3))
                if best_d is None or d < best_d:
                    best_d, best = d, face
            if best is not None:
                chosen = [best]
        elif spec == "all":
            chosen = candidates
        else:
            raise RuntimeError("unknown --faces spec: %s (all | near:x,y,z)" % spec)
        count = 0
        for face in chosen:
            if face.Select2(True, mark):
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
        all_bodies = self.bodies(m)
        solids = [b for b in all_bodies if self.body_kind(b) == "solid"]
        surfaces = [b for b in all_bodies if self.body_kind(b) == "surface"]
        out["bodies"] = len(all_bodies)
        out["solid_bodies"] = len(solids)
        out["surface_bodies"] = len(surfaces)
        if surfaces:
            out["surface_area_mm2"] = self.surface_area_mm2(m, "surface")
        if solids:
            mp = solids[0].GetMassProperties(1000.0)
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
    errors = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    warnings = win32.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
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


def cmd_plane_offset(b, a):
    return b.plane_offset(a.base, a.distance)


def cmd_surface_plane(b, a):
    return b.surface_plane(sketch=a.sketch)


def cmd_surface_extrude(b, a):
    return b.surface_extrude(a.depth, flip=a.flip, sketch=a.sketch)


def cmd_surface_revolve(b, a):
    return b.surface_revolve(a.angle, sketch=a.sketch, flip=a.flip)


def cmd_surface_loft(b, a):
    names = [x.strip() for x in a.sketches.split(",") if x.strip()]
    return b.surface_loft(names)


def cmd_surface_sweep(b, a):
    return b.surface_sweep(a.outline, a.path)


def cmd_thicken(b, a):
    return b.thicken(a.thickness, faces=a.faces, merge=not a.no_merge)


def cmd_offset_surface(b, a):
    return b.offset_surface(a.distance, faces=a.faces)


def cmd_sew(b, a):
    return b.sew(bodies=a.bodies, tolerance=a.tolerance)


def cmd_extend_surface(b, a):
    return b.extend_surface(a.distance, edges=a.edges)


def cmd_combine(b, a):
    return b.combine(a.operation, target=a.target, tools=a.tools)


def cmd_trim_surface(b, a):
    return b.trim_surface(tool=a.tool, target=a.target, mode=a.mode,
                          sew=a.sew, tool_mark=a.tool_mark,
                          target_mark=a.target_mark)


def cmd_fill_surface(b, a):
    return b.fill_surface(edges=a.edges, resolution=a.resolution)


def cmd_boundary_surface(b, a):
    sketches = [x.strip() for x in a.sketches.split(",") if x.strip()] \
        if a.sketches else None
    return b.boundary_surface(sketches=sketches, from_feature=a.from_feature)


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

    po = sub.add_parser("plane-offset")
    po.add_argument("--base", default="front")
    po.add_argument("--distance", type=float, required=True)

    sp = sub.add_parser("surface-plane")
    sp.add_argument("--sketch", default=None)

    sx = sub.add_parser("surface-extrude")
    sx.add_argument("--depth", type=float, required=True)
    sx.add_argument("--sketch", default=None)
    sx.add_argument("--flip", action="store_true")

    sr = sub.add_parser("surface-revolve")
    sr.add_argument("--angle", type=float, default=360.0)
    sr.add_argument("--sketch", default=None)
    sr.add_argument("--flip", action="store_true")

    sl = sub.add_parser("surface-loft")
    sl.add_argument("--sketches", required=True)

    swp = sub.add_parser("surface-sweep")
    swp.add_argument("--outline", required=True)
    swp.add_argument("--path", required=True)

    th = sub.add_parser("thicken")
    th.add_argument("--thickness", type=float, required=True)
    th.add_argument("--faces", default="all")
    th.add_argument("--no-merge", action="store_true")

    ofs = sub.add_parser("offset-surface")
    ofs.add_argument("--distance", type=float, required=True)
    ofs.add_argument("--faces", default="all")

    sew = sub.add_parser("sew")
    sew.add_argument("--bodies", default="surface", choices=["surface", "solid", "all"])
    sew.add_argument("--tolerance", type=float, default=0.01)

    ext = sub.add_parser("extend-surface")
    ext.add_argument("--distance", type=float, required=True)
    ext.add_argument("--edges", default="all")

    cb = sub.add_parser("combine")
    cb.add_argument("--operation", default="subtract",
                    choices=["add", "subtract", "intersect"])
    cb.add_argument("--target", default="largest")
    cb.add_argument("--tools", default="all")

    tse = sub.add_parser("trim-surface")
    tse.add_argument("--tool", default="last")
    tse.add_argument("--target", default="first")
    tse.add_argument("--mode", default="standard", choices=["standard", "mutual"])
    tse.add_argument("--sew", action="store_true")
    tse.add_argument("--tool-mark", type=int, default=1)
    tse.add_argument("--target-mark", type=int, default=4)

    fsv = sub.add_parser("fill-surface")
    fsv.add_argument("--edges", default="all")
    fsv.add_argument("--resolution", type=int, default=1)

    bfs = sub.add_parser("boundary-surface")
    bfs.add_argument("--sketches", default=None)
    bfs.add_argument("--from-feature", default=None)

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
    "plane-offset": cmd_plane_offset,
    "surface-plane": cmd_surface_plane,
    "surface-extrude": cmd_surface_extrude,
    "surface-revolve": cmd_surface_revolve,
    "surface-loft": cmd_surface_loft,
    "surface-sweep": cmd_surface_sweep,
    "thicken": cmd_thicken,
    "offset-surface": cmd_offset_surface,
    "sew": cmd_sew,
    "extend-surface": cmd_extend_surface,
    "combine": cmd_combine,
    "trim-surface": cmd_trim_surface,
    "fill-surface": cmd_fill_surface,
    "boundary-surface": cmd_boundary_surface,
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
