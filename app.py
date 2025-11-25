
# dxf_nesting_streamlit.py
# Single-file Streamlit app that uses your original nesting core (kept intact)
# UI only: Streamlit front-end, safer preview rendering, ZIP export.
# Run: pip install streamlit ezdxf pillow numpy
#       streamlit run dxf_nesting_streamlit.py

import math
import sys
import os
import io
import tempfile
import zipfile
import traceback
import uuid
import random
import json
import datetime
from collections import defaultdict
import numpy as np
import ezdxf
from PIL import Image, ImageDraw
import streamlit as st  # type: ignore
import base64
try:
    import pandas as pd  # for Excel export
except ImportError:
    pd = None

# Optional PDF report dependencies (ReportLab)
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (
        SimpleDocTemplate,
        Table,
        TableStyle,
        Paragraph,
        Spacer,
        Image as RLImage,
        PageBreak,
        KeepTogether,
    )
    from reportlab.lib.utils import ImageReader
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

DEFAULT_LOGO_BASE64 = None
DEFAULT_EXTERNAL_LOGO_PATH = "/mnt/data/.png"
# ---------------------------
# ---------- YOUR CORE (unchanged logic)
# ---------------------------
# CONFIG
DEFAULT_SHEET_W = 1000.0
DEFAULT_SHEET_H = 1000.0
DEFAULT_SPACING = 5.0
DEFAULT_KERF = 0.0
DEFAULT_ROT_STEP = 15
DEFAULT_QTY = 1
PROJECTS_ROOT = "projects"
DEFAULT_PROJECT_DISPLAY_NAME = " Unnamed file"

try:
    import shutil  # for project deletion
except Exception:
    shutil = None

# ---- Project metadata helpers ----
def _project_meta_path(pid: str):
    return os.path.join(PROJECTS_ROOT, pid, "meta.json")

def _save_project_meta():
    pid = st.session_state.get("project_id")
    if not pid:
        return
    meta = {
        "project_id": pid,
        "name": _normalized_project_name(st.session_state.get("project_name"), pid),
        "created_at": st.session_state.get("project_created_at"),
        "updated_at": datetime.datetime.now().isoformat(),
    }
    try:
        with open(_project_meta_path(pid), "w", encoding="utf-8") as fh:
            json.dump(meta, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _load_project_meta(pid: str):
    try:
        with open(_project_meta_path(pid), "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return None

def _normalized_project_name(name, pid=None):
    raw = name if isinstance(name, str) else ("" if name is None else str(name))
    trimmed = raw.strip()
    if not trimmed:
        return DEFAULT_PROJECT_DISPLAY_NAME
    if pid and trimmed == pid:
        return DEFAULT_PROJECT_DISPLAY_NAME
    return raw

def _create_new_project(initial_name: str = None):
    _ensure_projects_root()
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    pid = f"proj_{ts}_{uuid.uuid4().hex[:6]}"
    path = os.path.join(PROJECTS_ROOT, pid)
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass
    st.session_state.project_id = pid
    st.session_state.project_path = path
    st.session_state.project_created_at = datetime.datetime.now().isoformat()
    st.session_state.project_name = _normalized_project_name(initial_name or DEFAULT_PROJECT_DISPLAY_NAME, pid)
    # RESET ALL PROJECT-SPECIFIC STATE
    st.session_state.uploaded_meta = []
    st.session_state.nested_sheets = None
    st.session_state.part_summary = []
    st.session_state.config = {}
    st.session_state.config_inputs = {
        "sheet_w": DEFAULT_SHEET_W,
        "sheet_h": DEFAULT_SHEET_H,
        "spacing": DEFAULT_SPACING,
        "kerf": DEFAULT_KERF,
        "rotation_mode": "Free Rotation (Optimized)",
        "rotation_step": DEFAULT_ROT_STEP,
        "preview_canvas_px": 900,
        "show_bbox": True,
        "show_outline": True,
    }
    st.session_state.uploader_snapshot = None
    st.session_state.ignore_upload_names = set()
    st.session_state.uploader_key = f"uploader_{uuid.uuid4()}"
    st.session_state.stage = "SET"
    st.session_state.sheet_nav_number = 1
    st.session_state.export_sheet_number = 1
    _save_project_meta()
    _auto_save_project()  # initial snapshot
    return pid

def _delete_project(pid: str):
    if not pid or pid == "" or pid is None:
        return False
    path = os.path.join(PROJECTS_ROOT, pid)
    if not os.path.isdir(path):
        return False
    if shutil is None:
        return False
    try:
        shutil.rmtree(path)
        return True
    except Exception:
        return False

def _ensure_projects_root():
    try:
        os.makedirs(PROJECTS_ROOT, exist_ok=True)
    except Exception:
        pass

def _init_project_if_needed():
    _ensure_projects_root()
    if "project_id" not in st.session_state:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        pid = f"proj_{ts}_{uuid.uuid4().hex[:8]}"
        st.session_state.project_id = pid
        st.session_state.project_path = os.path.join(PROJECTS_ROOT, pid)
        st.session_state.project_created_at = datetime.datetime.now().isoformat()
        st.session_state.project_name = _normalized_project_name(DEFAULT_PROJECT_DISPLAY_NAME, pid)
        try:
            os.makedirs(st.session_state.project_path, exist_ok=True)
        except Exception:
            pass

def _auto_save_project(tag: str = "autosave"):
    path = st.session_state.get("project_path")
    if not path:
        return
    _ensure_projects_root()
    data = {
        "timestamp": datetime.datetime.now().isoformat(),
        "project_id": st.session_state.get("project_id"),
        "project_name": st.session_state.get("project_name"),
        "project_created_at": st.session_state.get("project_created_at"),
        "config_inputs": st.session_state.get("config_inputs"),
        "config": st.session_state.get("config"),
        "uploaded_meta": st.session_state.get("uploaded_meta"),
        "nested_sheets": st.session_state.get("nested_sheets"),
        "part_summary": st.session_state.get("part_summary"),
    }
    try:
        fname = os.path.join(path, f"{tag}.json")
        with open(fname, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _load_project(pid: str):
    root = os.path.join(PROJECTS_ROOT, pid)
    if not os.path.isdir(root):
        return False
    # prefer autosave.json else latest *.json
    target = os.path.join(root, "autosave.json")
    if not os.path.exists(target):
        jsons = [os.path.join(root, f) for f in os.listdir(root) if f.lower().endswith('.json')]
        if jsons:
            jsons.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            target = jsons[0]
        else:
            return False
    try:
        with open(target, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
        # restore
        st.session_state.project_id = payload.get("project_id", pid)
        st.session_state.project_path = root
        raw_name = payload.get("project_name") or ((_load_project_meta(pid) or {}).get("name"))
        st.session_state.project_name = _normalized_project_name(raw_name, st.session_state.project_id)
        st.session_state.project_created_at = payload.get("project_created_at") or ( (_load_project_meta(pid) or {}).get("created_at") )
        if payload.get("config_inputs"):
            st.session_state.config_inputs = payload["config_inputs"]
        if payload.get("config"):
            st.session_state.config = payload["config"]
        st.session_state.uploaded_meta = payload.get("uploaded_meta") or []
        st.session_state.nested_sheets = payload.get("nested_sheets")
        st.session_state.part_summary = payload.get("part_summary") or []
        st.session_state.stage = "CUT" if st.session_state.nested_sheets else "SET"
        return True
    except Exception:
        return False
FREE_ROTATION_SAMPLES = 36  # Number of angles to test for free rotation


def debug(msg):
    print("[DEBUG]", msg)


def point_in_polygon(point, polygon):
    x, y = point
    inside = False
    n = len(polygon)
    if n == 0:
        return False
    px1, py1 = polygon[0]
    for i in range(n + 1):
        px2, py2 = polygon[i % n]
        if min(py1, py2) < y <= max(py1, py2):
            denom = (py2 - py1) if abs(py2 - py1) > 1e-12 else 1e-12
            xinters = (px2 - px1) * (y - py1) / denom + px1
            if x < xinters:
                inside = not inside
        px1, py1 = px2, py2
    return inside


def polygon_centroid(pts):
    x_sum = sum(p[0] for p in pts)
    y_sum = sum(p[1] for p in pts)
    n = len(pts) or 1
    return (x_sum / n, y_sum / n)


def rotate_point_around(origin, pt, angle_deg):
    rad = math.radians(angle_deg)
    cos_r = math.cos(rad)
    sin_r = math.sin(rad)
    lx = pt[0] - origin[0]
    ly = pt[1] - origin[1]
    x_rot = lx * cos_r - ly * sin_r
    y_rot = lx * sin_r + ly * cos_r
    return (x_rot + origin[0], y_rot + origin[1])


def bbox_from_entities(ents):
    xs = []
    ys = []
    for ent in ents:
        t = ent["type"]
        if t == "LINE":
            xs += [ent["start"][0], ent["end"][0]]
            ys += [ent["start"][1], ent["end"][1]]
        elif t == "LWPOLYLINE":
            xs += [p[0] for p in ent["points"]]
            ys += [p[1] for p in ent["points"]]
        elif t in ("CIRCLE", "ARC"):
            c = ent["center"]
            r = ent["radius"]
            xs += [c[0] - r, c[0] + r]
            ys += [c[1] - r, c[1] + r]
        elif t == "SPLINE":
            xs += [p[0] for p in ent["points"]]
            ys += [p[1] for p in ent["points"]]
    if not xs or not ys:
        return (0.0, 0.0, 0.0, 0.0)
    return (min(xs), min(ys), max(xs), max(ys))


def get_part_centroid(part):
    """Calculate the geometric centroid of a part for better rotation origin"""
    all_points = []
    for ent in part.get("entities", []):
        if ent["type"] == "LWPOLYLINE":
            all_points.extend(ent["points"])
        elif ent["type"] in ("CIRCLE", "ARC"):
            all_points.append(ent["center"])
        elif ent["type"] == "LINE":
            all_points.extend([ent["start"], ent["end"]])
        elif ent["type"] == "SPLINE":
            all_points.extend(ent["points"])

    if not all_points:
        return (0, 0)
    return polygon_centroid(all_points)


def calculate_bbox_area(bbox):
    """Calculate area of a bounding box"""
    return (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])


def transform_entity_representation(ent):
    t = ent.dxftype()
    if t == "LINE":
        return {
            "type": "LINE",
            "start": (ent.dxf.start.x, ent.dxf.start.y),
            "end": (ent.dxf.end.x, ent.dxf.end.y),
            "layer": ent.dxf.layer,
        }
    if t == "LWPOLYLINE":
        pts_raw = list(ent.get_points("xyb"))
        pts = [(p[0], p[1]) for p in pts_raw]
        bulges = [p[2] for p in pts_raw]
        closed = bool(
            getattr(ent, "is_closed", ent.closed if hasattr(ent, "closed") else False)
        )
        return {
            "type": "LWPOLYLINE",
            "points": pts,
            "bulges": bulges,
            "closed": closed,
            "layer": ent.dxf.layer,
        }

    if t == "POLYLINE":
        try:
            pts = [(v.dxf.x, v.dxf.y) for v in ent.vertices()]
        except Exception:
            pts = []
        closed = bool(getattr(ent, "is_closed", False))
        return {
            "type": "LWPOLYLINE",
            "points": pts,
            "bulges": [0.0] * len(pts),
            "closed": closed,
            "layer": getattr(ent.dxf, "layer", ""),
        }

    if t == "CIRCLE":
        x, y, z = ent.dxf.center
        if ent.dxf.extrusion.z < 0:
            x = -x
        return {
            "type": "CIRCLE",
            "center": (x, y),
            "radius": ent.dxf.radius,
            "layer": ent.dxf.layer,
        }

    if t == "ARC":
        x = ent.dxf.center.x
        y = ent.dxf.center.y
        start_angle = ent.dxf.start_angle
        end_angle = ent.dxf.end_angle

        if ent.dxf.extrusion.z < 0:
            x = -x
            start_angle = (180 - start_angle) % 360
            end_angle = (180 - end_angle) % 360
            start_angle, end_angle = end_angle, start_angle
        return {
            "type": "ARC",
            "center": (x, y),
            "radius": ent.dxf.radius,
            "start_angle": start_angle,
            "end_angle": end_angle,
            "layer": ent.dxf.layer,
        }

    if t == "SPLINE":
        try:
            points = list(ent.flattening(0.01))
            pts = [(p[0], p[1]) for p in points]

            if hasattr(ent.dxf, "extrusion") and ent.dxf.extrusion.z < 0:
                pts = [(-p[0], p[1]) for p in pts]
            return {
                "type": "SPLINE",
                "points": pts,
                "layer": getattr(ent.dxf, "layer", ""),
            }
        except Exception as e:
            debug(f"Error processing SPLINE: {e}")
            return None

    if t == "INSERT":
        name = ent.dxf.name
        insert_point = (ent.dxf.insert.x, ent.dxf.insert.y)
        return {
            "type": "INSERT",
            "name": name,
            "insert": insert_point,
            "layer": getattr(ent.dxf, "layer", ""),
        }
    return None


def detect_entity_types(doc):
    counts = defaultdict(int)
    msp = doc.modelspace()
    for e in msp:
        counts[e.dxftype()] += 1
    debug("Entity counts: " + ", ".join(f"{k}:{v}" for k, v in counts.items()))
    return counts


def flatten_insert(doc, ent_insert):
    name = ent_insert["name"]
    if name not in doc.blocks:
        return []
    blk = doc.blocks.get(name)
    flattened = []
    for e in blk:
        rep = transform_entity_representation(e)
        if rep:
            if "start" in rep:
                rep["start"] = (
                    rep["start"][0] + ent_insert["insert"][0],
                    rep["start"][1] + ent_insert["insert"][1],
                )
            if "end" in rep:
                rep["end"] = (
                    rep["end"][0] + ent_insert["insert"][0],
                    rep["end"][1] + ent_insert["insert"][1],
                )
            if "points" in rep:
                rep["points"] = [
                    (p[0] + ent_insert["insert"][0], p[1] + ent_insert["insert"][1])
                    for p in rep["points"]
                ]
            if "center" in rep:
                rep["center"] = (
                    rep["center"][0] + ent_insert["insert"][0],
                    rep["center"][1] + ent_insert["insert"][1],
                )
            flattened.append(rep)
    return flattened


def group_entities_universal(entities, doc=None):
    outlines = [e for e in entities if e["type"] == "LWPOLYLINE" and e.get("closed")]
    if not outlines:
        outlines = [e for e in entities if e["type"] == "LWPOLYLINE"]
    inserts = [e for e in entities if e["type"] == "INSERT"]
    if outlines:
        parts = []
        others = [e for e in entities if e not in outlines]
        for out in outlines:
            part = {
                "outer": out,
                "holes": [],
                "circles": [],
                "arcs": [],
                "splines": [],
                "others": [],
            }
            poly = out.get("points", [])
            for h in others:
                if h["type"] == "CIRCLE":
                    if poly and point_in_polygon(h["center"], poly):
                        part["circles"].append(h)
                    else:
                        part["others"].append(h)
                elif h["type"] == "LWPOLYLINE":
                    if poly:
                        c = polygon_centroid(h["points"])
                        if point_in_polygon(c, poly):
                            part["holes"].append(h)
                        else:
                            part["others"].append(h)
                    else:
                        part["others"].append(h)
                elif h["type"] == "ARC":
                    if poly and point_in_polygon(h["center"], poly):
                        part["arcs"].append(h)
                    else:
                        part["others"].append(h)
                elif h["type"] == "SPLINE":
                    if poly:
                        c = polygon_centroid(h["points"])
                        if point_in_polygon(c, poly):
                            part["splines"].append(h)
                        else:
                            part["others"].append(h)
                    else:
                        part["others"].append(h)
                else:
                    part["others"].append(h)
            parts.append(part)
        return parts
    if inserts and doc is not None:
        parts = []
        for ins in inserts:
            flat = flatten_insert(doc, ins)
            if flat:
                outlines_block = [
                    e for e in flat if e["type"] == "LWPOLYLINE" and e.get("closed")
                ]
                if outlines_block:
                    out = outlines_block[0]
                    others = [e for e in flat if e is not out]
                    part = {
                        "outer": out,
                        "holes": [e for e in others if e["type"] == "LWPOLYLINE"],
                        "circles": [e for e in others if e["type"] == "CIRCLE"],
                        "arcs": [e for e in others if e["type"] == "ARC"],
                        "splines": [e for e in others if e["type"] == "SPLINE"],
                        "others": [],
                    }
                    parts.append(part)
                else:
                    parts.append(
                        {
                            "outer": None,
                            "holes": flat,
                            "circles": [],
                            "arcs": [],
                            "splines": [],
                            "others": [],
                        }
                    )
        return parts
    return [
        {
            "outer": None,
            "holes": entities,
            "circles": [],
            "arcs": [],
            "splines": [],
            "others": [],
        }
    ]


def calculate_relative_offsets_for_part(part):
    outer = part.get("outer")
    if not outer:
        bbox = bbox_from_entities(part.get("holes", []))
        origin = (bbox[0], bbox[1])
    else:
        origin = (
            min(p[0] for p in outer.get("points", [(0, 0)])),
            min(p[1] for p in outer.get("points", [(0, 0)])),
        )
    part["origin"] = origin
    offsets = []
    for hole in part.get("holes", []):
        if hole["type"] == "LWPOLYLINE" and hole.get("points"):
            c = polygon_centroid(hole["points"])
            offsets.append((hole, (c[0] - origin[0], c[1] - origin[1])))
    for circ in part.get("circles", []):
        offsets.append(
            (circ, (circ["center"][0] - origin[0], circ["center"][1] - origin[1]))
        )
    for arc in part.get("arcs", []):
        offsets.append(
            (arc, (arc["center"][0] - origin[0], arc["center"][1] - origin[1]))
        )
    for spline in part.get("splines", []):
        if spline.get("points"):
            c = polygon_centroid(spline["points"])
            offsets.append((spline, (c[0] - origin[0], c[1] - origin[1])))
    part["relative_offsets"] = offsets


def transform_point(pt, origin, angle_deg, translate_x, translate_y):
    """Rotate a point around `origin` by angle_deg and then translate.

    This is used for both preview drawing and DXF export.
    """
    rad = math.radians(angle_deg)
    cos_r = math.cos(rad)
    sin_r = math.sin(rad)

    # Local coordinates around rotation origin
    local_x = pt[0] - origin[0]
    local_y = pt[1] - origin[1]

    # Correct 2D rotation
    x_rot = local_x * cos_r - local_y * sin_r
    y_rot = local_x * sin_r + local_y * cos_r

    # Translate back and apply final translation offset
    return (
        x_rot + origin[0] + translate_x,
        y_rot + origin[1] + translate_y,
    )


# ---------------------------
# ---------- PDF REPORT BUILDER
# ---------------------------
def _make_logo_image_from_text(text="Ei8", size_px=120, bg=(255, 255, 255), fg=(0, 0, 0)):
    img = Image.new("RGB", (size_px, size_px), bg)
    d = ImageDraw.Draw(img)
    # Simple centered text logo
    # Pillow default font is fine for placeholder
    tw, th = d.textlength(text), 20
    d.text(((size_px - tw) / 2, (size_px - th) / 2), text, fill=fg)
    return img


def _on_page(canvas, doc, logo_reader, title_text):
    canvas.saveState()
    width, height = doc.pagesize
    # Header band
    header_h = 20 * mm
    canvas.setFillColorRGB(0.96, 0.96, 0.96)
    canvas.rect(0, height - header_h, width, header_h, stroke=0, fill=1)
    # Logo
    if logo_reader is not None:
        logo_w = 18 * mm
        logo_h = 18 * mm
        canvas.drawImage(logo_reader, 10 * mm, height - header_h + 1 * mm, logo_w, logo_h, preserveAspectRatio=True, mask='auto')
    # Title
    canvas.setFillColorRGB(0, 0, 0)
    canvas.setFont("Helvetica-Bold", 13)
    canvas.drawString(32 * mm, height - 12 * mm, title_text)
    # Footer
    canvas.setFont("Helvetica", 9)
    canvas.setFillColorRGB(0.4, 0.4, 0.4)
    canvas.drawRightString(width - 10 * mm, 8 * mm, f"Page {doc.page}")
    canvas.restoreState()


def build_pdf_report(
    sheets,
    cfg,
    preview_canvas_px,
    show_bbox,
    show_outline,
    scope="Single sheet",
    sel_idx=1,
    uploaded_meta=None,
    config_inputs=None,
    logo_bytes: bytes | None = None,
    title_text: str = "Optimization Nesting Report",
):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab is not installed. Please install 'reportlab'.")

    buf = io.BytesIO()
    # Landscape A4 for wider nesting previews
    pagesize = landscape(A4)
    doc = SimpleDocTemplate(
        buf,
        pagesize=pagesize,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=25 * mm,
        bottomMargin=15 * mm,
        title=title_text,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SmallGray", fontSize=9, textColor=colors.grey))
    story = []

    # Prepare logo / cover image
    logo_data = None
    if logo_bytes:
        logo_data = logo_bytes
    else:
        # Try explicit external path first
        if os.path.exists(DEFAULT_EXTERNAL_LOGO_PATH):
            try:
                with open(DEFAULT_EXTERNAL_LOGO_PATH, "rb") as fh:
                    logo_data = fh.read()
            except Exception:
                logo_data = None
        if logo_data is None:
            cover_candidates = [
                "ei8_logo.png",
                "report_logo.png",
                "cover.png",
                os.path.join("assets", "ei8_logo.png"),
                os.path.join("assets", "report_logo.png"),
                os.path.join("assets", "cover.png"),
            ]
            for cand in cover_candidates:
                pth = os.path.join(os.getcwd(), cand)
                if os.path.exists(pth):
                    try:
                        with open(pth, "rb") as fh:
                            logo_data = fh.read()
                        break
                    except Exception:
                        pass
    if logo_data is None:
        try:
            logo_data = base64.b64decode(DEFAULT_LOGO_BASE64)
        except Exception:
            logo_data = None

    logo_image = None
    if logo_data is not None:
        try:
            logo_image = Image.open(io.BytesIO(logo_data)).convert("RGBA")
        except Exception:
            logo_image = None

    if logo_image is None:
        logo_image = _make_logo_image_from_text("Ei8")
        buf_tmp = io.BytesIO()
        logo_image.save(buf_tmp, format="PNG")
        logo_data = buf_tmp.getvalue()

    logo_canvas_reader = ImageReader(io.BytesIO(logo_data)) if logo_data else None

    # Collect which sheets to include
    target_sheets = sheets if scope == "All sheets" else [sheets[sel_idx - 1]]

    # ---------------- COVER PAGE (logo + meta side-by-side) -----------------
    styles.add(ParagraphStyle(name="SectionHeader", fontSize=11, leading=14, alignment=0, fontName="Helvetica-Bold", textColor=colors.HexColor('#0A0F1A')))
    title_para = Paragraph(f"<b>{title_text}</b>", styles["Title"]) if "Title" in styles else Paragraph(f"<b>{title_text}</b>", styles["Normal"])

    # Prepare logo scaled to fixed height for consistent alignment
    logo_flow = Spacer(1,1)
    if logo_image is not None:
        lw, lh = logo_image.size
        target_h = 40*mm
        scale = min(target_h / lh, 1.0)
        disp_w = lw * scale
        disp_h = lh * scale
        logo_buf = io.BytesIO()
        logo_image.save(logo_buf, format="PNG")
        logo_buf.seek(0)
        logo_flow = RLImage(logo_buf, width=disp_w, height=disp_h)

    rot_mode = cfg.get("rotation_mode", (config_inputs or {}).get("rotation_mode", "-"))
    rot_step = (config_inputs or {}).get("rotation_step", cfg.get("rot_step", DEFAULT_ROT_STEP))
    meta_rows = [
        ["Material", str(cfg.get("material", "-"))],
        ["Thickness", str(cfg.get("thickness", "-"))],
        ["Sheet Size", f"{cfg['sheet_w']} x {cfg['sheet_h']} mm"],
        ["Kerf", str(cfg.get("kerf", cfg.get("kerf_width", 0)))],
        ["Spacing", str(cfg.get("spacing", DEFAULT_SPACING))],
        ["Rotation Mode", str(rot_mode)],
        ["Rotation Step", str(rot_step)],
        ["Sheets", str(len(target_sheets))],
    ]
    meta_tbl = Table(meta_rows, colWidths=[35*mm, 55*mm])
    meta_tbl.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor('#d0d7e2')),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor('#f1f5f9')),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("ALIGN", (0,0), (0,-1), "RIGHT"),
    ]))

    # Aggregates (compute before building tables below)
    total_files = len([m for m in (uploaded_meta or []) if m.get("valid")])
    requested_parts = 0
    placed_parts = sum(len(sh) for sh in sheets)
    for m in (uploaded_meta or []):
        if m.get("valid"):
            requested_parts += int(m.get("qty", 1)) * len(m.get("parts", []))
    sheet_utils = []
    sheet_area = cfg["sheet_w"] * cfg["sheet_h"] if cfg.get("sheet_w") and cfg.get("sheet_h") else 0
    for sh in sheets:
        part_area_sum = sum(calculate_bbox_area(p.get("placed_bbox", (0,0,0,0))) for p in sh)
        util = (part_area_sum / sheet_area * 100.0) if sheet_area > 0 else 0.0
        sheet_utils.append(util)
    overall_avg_util = sum(sheet_utils) / len(sheet_utils) if sheet_utils else 0.0
    overall_min_util = min(sheet_utils) if sheet_utils else 0.0
    overall_max_util = max(sheet_utils) if sheet_utils else 0.0
    agg_rows = [
        ["Requested Parts", str(requested_parts)],
        ["Placed Parts", str(placed_parts)],
        ["Avg Util", f"{overall_avg_util:.2f}%"],
        ["Min Util", f"{overall_min_util:.2f}%"],
        ["Max Util", f"{overall_max_util:.2f}%"],
    ]
    agg_tbl = Table(agg_rows, colWidths=[30*mm, 60*mm])
    agg_tbl.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor('#d0d7e2')),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor('#f1f5f9')),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("ALIGN", (0,0), (0,-1), "RIGHT"),
    ]))

    # Simple cover layout: logo centered, title below, then meta & aggregates side-by-side
    # Logo
    story.append(logo_flow)
    story.append(Spacer(1, 10))
    story.append(title_para)
    story.append(Spacer(1, 14))
    # Two-column info tables
    info_table = Table(
        [
            [meta_tbl, agg_tbl]
        ],
        colWidths=[doc.width*0.55, doc.width*0.45]
    )
    info_table.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 2),
        ("RIGHTPADDING", (0,0), (-1,-1), 2),
        ("TOPPADDING", (0,0), (-1,-1), 0),
        ("BOTTOMPADDING", (0,0), (-1,-1), 0),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 16))

    # Part summary (single column) under info
    if sheets:
        part_counts = {}
        for sh in sheets:
            for part in sh:
                lbl = part.get("label") or part.get("source_file") or "Part"
                part_counts[lbl] = part_counts.get(lbl, 0) + 1
        if part_counts:
            pc_rows = [["Part", "Qty"]] + [[k, str(v)] for k, v in sorted(part_counts.items(), key=lambda x: x[0].lower())]
            pc_tbl = Table(pc_rows, repeatRows=1, colWidths=[doc.width*0.6, doc.width*0.4])
            pc_tbl.setStyle(TableStyle([
                ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor('#d0d7e2')),
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor('#e2e8f0')),
                ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
                ("FONTSIZE", (0,0), (-1,-1), 8),
            ]))
            story.append(Paragraph("<b>Part Summary</b>", styles["SectionHeader"]))
            story.append(pc_tbl)
            story.append(Spacer(1, 10))

    # Uploaded files table (compact) below summary
    if uploaded_meta:
        file_rows = [["#", "Label", "Filename", "Qty", "H(mm)", "W(mm)"]]
        for i, m in enumerate(uploaded_meta, start=1):
            if not m.get("valid"):
                continue
            parts_list = m.get("parts", []) or []
            b_w = b_h = 0.0
            if parts_list:
                p0 = parts_list[0]
                bb = p0.get("orig_bbox") or bbox_from_entities(p0.get("entities", []))
                b_w = max(0.0, (bb[2] - bb[0]))
                b_h = max(0.0, (bb[3] - bb[1]))
            file_rows.append([
                i,
                str(m.get("label", m.get("name", "")))[:20],
                str(m.get("name", ""))[:25],
                str(int(m.get("qty", 1))),
                f"{b_h:.1f}",
                f"{b_w:.1f}",
            ])
        if len(file_rows) > 1:
            files_tbl = Table(file_rows, repeatRows=1, colWidths=[8*mm, 38*mm, 56*mm, 12*mm, 18*mm, 18*mm])
            files_tbl.setStyle(TableStyle([
                ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor('#d0d7e2')),
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor('#e2e8f0')),
                ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
                ("FONTSIZE", (0,0), (-1,-1), 7),
            ]))
            story.append(Paragraph("<b>Uploaded Files</b>", styles["SectionHeader"]))
            story.append(files_tbl)
            story.append(Spacer(1, 12))

    story.append(PageBreak())

    # ---------------- SHEET PAGES ------------------------------------------
    rot_mode = cfg.get("rotation_mode", (config_inputs or {}).get("rotation_mode", "-"))
    rot_step = (config_inputs or {}).get("rotation_step", cfg.get("rot_step", DEFAULT_ROT_STEP))
    meta_rows = [
        ["Material", str(cfg.get("material", "-"))],
        ["Thickness", str(cfg.get("thickness", "-"))],
        ["Sheet Size", f"{cfg['sheet_w']} x {cfg['sheet_h']} mm"] ,
        ["Kerf", str(cfg.get("kerf", cfg.get("kerf_width", 0)))],
        ["Spacing", str(cfg.get("spacing", DEFAULT_SPACING))],
        ["Rotation Mode", str(rot_mode)],
        ["Rotation Step", str(rot_step)],
        ["Sheets", str(len(target_sheets))],
    ]
    meta_tbl = Table(meta_rows, colWidths=[40*mm, doc.width - 40*mm])
    meta_tbl.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("ALIGN", (0,0), (0,-1), "RIGHT"),
    ]))

    # We'll add meta first; header drawn via onPage
    story.append(Spacer(1, 8))
    story.append(meta_tbl)
    story.append(Spacer(1, 8))

    # ---- GLOBAL AGGREGATES --------------------------------------------------
    total_files = len([m for m in (uploaded_meta or []) if m.get("valid")])
    requested_parts = 0
    placed_parts = sum(len(sh) for sh in sheets)
    for m in (uploaded_meta or []):
        if m.get("valid"):
            requested_parts += int(m.get("qty", 1)) * len(m.get("parts", []))
    sheet_utils = []
    sheet_area = cfg["sheet_w"] * cfg["sheet_h"] if cfg.get("sheet_w") and cfg.get("sheet_h") else 0
    for sh in sheets:
        part_area_sum = sum(calculate_bbox_area(p.get("placed_bbox", (0,0,0,0))) for p in sh)
        util = (part_area_sum / sheet_area * 100.0) if sheet_area > 0 else 0.0
        sheet_utils.append(util)
    overall_avg_util = sum(sheet_utils) / len(sheet_utils) if sheet_utils else 0.0
    overall_min_util = min(sheet_utils) if sheet_utils else 0.0
    overall_max_util = max(sheet_utils) if sheet_utils else 0.0
    agg_rows = [
        ["Total Files", str(total_files)],
        ["Requested Parts", str(requested_parts)],
        ["Placed Parts", str(placed_parts)],
        ["Avg Utilization", f"{overall_avg_util:.2f}%"],
        ["Min Utilization", f"{overall_min_util:.2f}%"],
        ["Max Utilization", f"{overall_max_util:.2f}%"],
    ]
    agg_tbl = Table(agg_rows, colWidths=[40*mm, doc.width - 40*mm])
    agg_tbl.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("ALIGN", (0,0), (0,-1), "RIGHT"),
    ]))
    story.append(agg_tbl)
    story.append(Spacer(1, 8))

    # Aggregated part quantities
    if sheets:
        part_counts = {}
        for sh in sheets:
            for part in sh:
                lbl = part.get("label") or part.get("source_file") or "Part"
                part_counts[lbl] = part_counts.get(lbl, 0) + 1
        if part_counts:
            pc_rows = [["Part", "Quantity"]] + [[k, str(v)] for k, v in sorted(part_counts.items(), key=lambda x: x[0].lower())]
            pc_tbl = Table(pc_rows, repeatRows=1, colWidths=[doc.width/2.0, doc.width/2.0])
            pc_tbl.setStyle(TableStyle([
                ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
                ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
                ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
                ("FONTSIZE", (0,0), (-1,-1), 8),
            ]))
            story.append(Paragraph("<b>Part Summary</b>", styles["Normal"]))
            story.append(pc_tbl)
            story.append(Spacer(1, 8))

    # Uploaded files table (if available)
    if uploaded_meta:
        file_rows = [["#", "Label", "Filename", "Qty", "Height (mm)", "Width (mm)", "Entities"]]
        for i, m in enumerate(uploaded_meta, start=1):
            if not m.get("valid"):
                continue
            parts = m.get("parts", []) or []
            b_w = b_h = 0.0
            if parts:
                p0 = parts[0]
                bb = p0.get("orig_bbox") or bbox_from_entities(p0.get("entities", []))
                b_w = max(0.0, (bb[2] - bb[0]))
                b_h = max(0.0, (bb[3] - bb[1]))
            ent_counts = m.get("entity_counts", {})
            ent_str = ", ".join(f"{k}:{v}" for k, v in sorted(ent_counts.items())) if ent_counts else "-"
            file_rows.append([
                i,
                str(m.get("label", m.get("name", ""))),
                str(m.get("name", "")),
                str(int(m.get("qty", 1))),
                f"{b_h:.1f}",
                f"{b_w:.1f}",
                ent_str,
            ])
        if len(file_rows) > 1:
            story.append(Paragraph("<b>Uploaded Files</b>", styles["Normal"]))
            files_tbl = Table(file_rows, repeatRows=1, colWidths=[8*mm, 35*mm, 50*mm, 12*mm, 20*mm, 20*mm, doc.width - (8+35+50+12+20+20)*mm])
            files_tbl.setStyle(TableStyle([
                ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
                ("FONTSIZE", (0,0), (-1,-1), 8),
                ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
                ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
            ]))
            story.append(files_tbl)
            story.append(Spacer(1, 6))

        # Thumbnail gallery (limit first 12 files)
        thumbs = []
        for m in uploaded_meta[:12]:
            if not m.get("valid"):
                continue
            thumb_img = create_file_preview_thumbnail(m, width=160, height=100)
            b = io.BytesIO()
            thumb_img.save(b, format="PNG")
            b.seek(0)
            thumbs.append(RLImage(b, width=55*mm, height=34*mm))
        if thumbs:
            story.append(Paragraph("<b>File Previews</b>", styles["Normal"]))
            row = []
            grid = []
            for t in thumbs:
                row.append(t)
                if len(row) == 4:  # 4 per row for tighter fit
                    grid.append(row)
                    row = []
            if row:
                grid.append(row + [Spacer(1,0)]*(4-len(row)))
            gal_tbl = Table(grid, colWidths=[doc.width/4.0]*4)
            gal_tbl.setStyle(TableStyle([
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ]))
            story.append(gal_tbl)
            story.append(Spacer(1, 8))

    # Cover page ends, ensure next page starts sheets
    story.append(PageBreak())

    # Per-sheet pages
    for s_idx, sh in enumerate(target_sheets, start=(1 if scope == "All sheets" else sel_idx)):
        # Compute utilization
        sheet_area = cfg["sheet_w"] * cfg["sheet_h"]
        part_area_sum = sum(calculate_bbox_area(p.get("placed_bbox", (0,0,0,0))) for p in sh)
        utilization = (part_area_sum / sheet_area * 100.0) if sheet_area > 0 else 0.0

        # Small sheet header table
        hdr_rows = [
            [Paragraph(f"<b>Sheet {s_idx}</b>", styles["Normal"]), f"Utilization: {utilization:.2f}%", f"Parts: {len(sh)}"],
        ]
        hdr = Table(hdr_rows, colWidths=[65*mm, 55*mm, doc.width - (65+55)*mm])
        hdr.setStyle(TableStyle([
            ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
            ("FONTSIZE", (0,0), (-1,-1), 10),
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor('#f1f5f9')),
            ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor('#d0d7e2')),
        ]))
        story.append(hdr)
        story.append(Spacer(1, 6))

        # Render preview image using existing preview function
        img = create_sheet_preview_image(
            sh,
            cfg["sheet_w"],
            cfg["sheet_h"],
            canvas_px=preview_canvas_px,
            draw_bbox=show_bbox,
            draw_outline=show_outline,
        )
        if img is not None:
            if img.mode != "RGB":
                img = img.convert("RGB")
            # Scale image to fit both width and available height
            w, h = img.size
            max_width = doc.width
            # Reserve some vertical space for header + table; allow at most 55% of frame height for preview
            max_height = doc.height * 0.45
            scale = min(max_width / w, max_height / h, 1.0)
            disp_w = max(1, w * scale)
            disp_h = max(1, h * scale)
            # If still larger than frame, cap further
            if disp_h > max_height:
                scale2 = max_height / disp_h
                disp_w *= scale2
                disp_h = max_height
            img_b = io.BytesIO()
            img.save(img_b, format="PNG")
            img_b.seek(0)
            try:
                story.append(RLImage(img_b, width=disp_w, height=disp_h))
            except Exception:
                # Fallback: shrink aggressively
                fallback_scale = 0.4
                story.append(RLImage(img_b, width=w * fallback_scale, height=h * fallback_scale))
        else:
            story.append(Paragraph("Preview not available", styles["SmallGray"]))

        # Parts table (detailed)
        if sh:
            # Section header to mimic screenshot style
            story.append(Spacer(1, 4))
            story.append(Paragraph("<b>Parts</b>", styles["SectionHeader"]))
            story.append(Spacer(1, 6))

            def _part_thumb(bbox):
                """Create a simple proportional rectangle thumbnail representing bounding box."""
                w = max(1, bbox[2] - bbox[0])
                h = max(1, bbox[3] - bbox[1])
                # Normalize to 120x120 canvas keeping aspect
                canvas_sz = 120
                scale = min(canvas_sz / w, canvas_sz / h)
                disp_w = int(w * scale)
                disp_h = int(h * scale)
                img = Image.new("RGB", (canvas_sz, canvas_sz), (236, 240, 244))  # light gray bg
                d = ImageDraw.Draw(img)
                left = (canvas_sz - disp_w) // 2
                top = (canvas_sz - disp_h) // 2
                d.rectangle([left, top, left + disp_w, top + disp_h], fill=(180, 192, 200))
                # Optional id centered
                return img

            for i, part in enumerate(sh, start=1):
                bbox = part.get("placed_bbox", (0,0,0,0))
                pw_mm = max(0.0, bbox[2] - bbox[0])
                ph_mm = max(0.0, bbox[3] - bbox[1])
                pw_cm = pw_mm / 10.0
                ph_cm = ph_mm / 10.0
                perimeter_mm = 2 * (pw_mm + ph_mm)
                perimeter_m = perimeter_mm / 1000.0
                area_cm2 = (pw_cm * ph_cm)
                # External contour assumed 1; internal contours approximate from holes + circles
                internal_cnt = len(part.get("holes", [])) + len(part.get("circles", [])) + len(part.get("arcs", []))
                external_cnt = 1
                name_val = part.get("label") or part.get("source_file") or f"Part {i}"
                # Thumbnail image
                thumb_img = _part_thumb(bbox)
                b_io = io.BytesIO()
                thumb_img.save(b_io, format="PNG")
                b_io.seek(0)
                thumb = RLImage(b_io, width=30*mm, height=30*mm)  # scale image cell

                # Attribute rows (two column label/value except first row three cells like screenshot)
                attr_rows = [
                    [f"ID {part.get('group_id', i)}", "Name", name_val[:40]],
                    ["Nested Quantity", "1 / 1"],
                    ["Bounding box (W x H)", f"{pw_cm:.3f} cm x {ph_cm:.3f} cm"],
                    ["Cut length / mark length", f"{perimeter_m:.3f} m / 0.000 mm"],
                    ["Surface", f"{area_cm2:.2f} cm²"],
                    ["External / Internal contours", f"{external_cnt} / {internal_cnt}"],
                ]
                # Determine column widths for attribute table
                # First row has three columns; we will allow dynamic merge for consistency
                # Build table data ensuring rectangular shape: convert 2-col rows into 3-cols with span indicator
                table_data = []
                for ridx, r in enumerate(attr_rows):
                    if ridx == 0:
                        table_data.append(r)
                    else:
                        # r has 2 columns -> expand to 2 with span over third
                        table_data.append([r[0], r[1], ""])
                # Widen first column to reduce wrapping; second narrower (value spans 2 & 3)
                attr_tbl = Table(table_data, colWidths=[60*mm, 35*mm, 50*mm])
                ts = [
                    ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
                    ("FONTSIZE", (0,0), (-1,-1), 8),
                    ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor('#d0d7e2')),
                    ("BACKGROUND", (0,0), (-1,0), colors.HexColor('#f1f5f9')),
                    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ]
                # Spans for rows other than first: value text sits in column 1 spanning columns 1+2
                for ridx in range(1, len(table_data)):
                    ts.append(("SPAN", (1, ridx), (2, ridx)))
                # Slight smaller font for long label rows (Bounding box & contour row)
                ts.append(("FONTSIZE", (0,2), (0,2), 7))
                ts.append(("FONTSIZE", (0,5), (0,5), 7))
                attr_tbl.setStyle(TableStyle(ts))

                card_tbl = Table([[thumb, attr_tbl]], colWidths=[32*mm, doc.width - 32*mm])
                card_tbl.setStyle(TableStyle([
                    ("VALIGN", (0,0), (-1,-1), "TOP"),
                    ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor('#d0d7e2')),
                    ("BACKGROUND", (0,0), (0,0), colors.HexColor('#eef2f6')),
                ]))
                story.append(card_tbl)
                story.append(Spacer(1, 8))

        if (scope == "All sheets") or (s_idx != sel_idx):
            story.append(PageBreak())

    # Build with custom header/footer
    def _on_first_page(canv, _doc):
        _on_page(canv, _doc, logo_canvas_reader, title_text)

    def _on_later_pages(canv, _doc):
        _on_page(canv, _doc, logo_canvas_reader, title_text)

    doc.build(story, onFirstPage=_on_first_page, onLaterPages=_on_later_pages)
    buf.seek(0)
    return buf.getvalue()


def transform_entity(ent, origin, ang, tx, ty):
    t = ent["type"]
    if t == "LINE":
        return {
            "type": "LINE",
            "start": transform_point(ent["start"], origin, ang, tx, ty),
            "end": transform_point(ent["end"], origin, ang, tx, ty),
            "layer": ent.get("layer"),
        }
    if t == "LWPOLYLINE":
        pts = [transform_point(p, origin, ang, tx, ty) for p in ent["points"]]
        return {
            "type": "LWPOLYLINE",
            "points": pts,
            "bulges": ent.get("bulges", []),
            "closed": ent.get("closed", False),
            "layer": ent.get("layer"),
        }
    if t == "CIRCLE":
        c = transform_point(ent["center"], origin, ang, tx, ty)
        return {
            "type": "CIRCLE",
            "center": c,
            "radius": ent["radius"],
            "layer": ent.get("layer"),
        }
    if t == "ARC":
        c = transform_point(ent["center"], origin, ang, tx, ty)
        new_start = (ent["start_angle"] + ang) % 360
        new_end = (ent["end_angle"] + ang) % 360
        return {
            "type": "ARC",
            "center": c,
            "radius": ent["radius"],
            "start_angle": new_start,
            "end_angle": new_end,
            "layer": ent.get("layer"),
        }
    if t == "SPLINE":
        pts = [transform_point(p, origin, ang, tx, ty) for p in ent["points"]]
        return {"type": "SPLINE", "points": pts, "layer": ent.get("layer")}
    return None


def boxes_intersect(a, b):
    return not (a[2] <= b[0] or a[0] >= b[2] or a[3] <= b[1] or a[1] >= b[3])


def find_best_rotation_for_position(
    part, x, y, placed_bboxes, sheet_w, sheet_h, spacing, rotations
):
    origin = part.get("centroid", part.get("origin", (0, 0)))
    entities = part["entities"]

    best_rotation = None
    best_bbox = None
    best_area = float("inf")

    for ang in rotations:
        rot_bbox = bbox_from_entities(
            [transform_entity(e, origin, ang, 0, 0) for e in entities]
        )
        bbox_w = rot_bbox[2] - rot_bbox[0]
        bbox_h = rot_bbox[3] - rot_bbox[1]

        if x + bbox_w + spacing > sheet_w or y + bbox_h + spacing > sheet_h:
            continue

        candidate = (x, y, x + bbox_w, y + bbox_h)

        if all(not boxes_intersect(candidate, pb) for pb in placed_bboxes):
            area = bbox_w * bbox_h
            if area < best_area:
                best_area = area
                best_rotation = ang
                best_bbox = (candidate, (x, y), rot_bbox)

    return best_rotation, best_bbox


def bottom_left_place_with_rotation(
    part, placed_bboxes, sheet_w, sheet_h, spacing, rotations, aggressive=False
):
    """
    Try to place `part` on the sheet using a bottom‑left heuristic.

    - Candidate X/Y positions come from edges of already placed parts.
    - When `aggressive` is True we also use left/bottom edges as anchors.
    - We do NOT cap the number of candidate positions; this lets us exploit
      narrow leftover strips and corners.
    """
    potential_xs = {spacing}
    potential_ys = {spacing}

    for pb in placed_bboxes:
        # Right / top edges
        potential_xs.add(pb[2] + spacing)
        potential_ys.add(pb[3] + spacing)
        if aggressive:
            # Also consider left / bottom edges for extra anchors
            potential_xs.add(pb[0] + spacing)
            potential_ys.add(pb[1] + spacing)

    sorted_xs = sorted(potential_xs)
    sorted_ys = sorted(potential_ys)

    best_placement = None
    best_score = None

    for y in sorted_ys:
        for x in sorted_xs:
            rotation, bbox_info = find_best_rotation_for_position(
                part, x, y, placed_bboxes, sheet_w, sheet_h, spacing, rotations
            )
            if rotation is None or bbox_info is None:
                continue

            candidate, anchor, rot_bbox = bbox_info

            # Score: lower Y (bottom), then lower X (left), then smaller bbox area.
            area = calculate_bbox_area(candidate)
            score = (candidate[3], candidate[0], area)

            if best_score is None or score < best_score:
                best_score = score
                best_placement = (rotation, candidate, anchor, rot_bbox)

    return best_placement

def nest_parts_improved(
    parts,
    sheet_w,
    sheet_h,
    spacing,
    kerf,
    rotations,
    free_rotation=False,
    advanced_sort=True,
    enable_compaction=True,
    rotation_prune_tolerance=0.05,
    aggressive_packing=True,
):
    """
    Main nesting routine with:
    - precomputation of geometry
    - bottom‑left placement with rotation
    - compaction
    - cross‑sheet relocation
    - global repack (only accepted if *all* parts are kept and sheet count
      does not increase)

    IMPORTANT: upload order is ignored. Parts are shuffled and then ordered
    by geometry, not by file sequence.
    """

    # ---- PREPARE PART GEOMETRY ------------------------------------------------
    # Randomize initial order so we don't depend on upload order
    random.shuffle(parts)

    for part in parts:
        ents = []
        if part.get("outer"):
            ents.append(part["outer"])
        ents += (
            part.get("holes", [])
            + part.get("circles", [])
            + part.get("arcs", [])
            + part.get("splines", [])
            + part.get("others", [])
        )
        part["entities"] = ents
        part["orig_bbox"] = bbox_from_entities(ents)
        part["centroid"] = get_part_centroid(part)
        calculate_relative_offsets_for_part(part)

    # ---- ORDERING: LARGE / SHEET‑LIKE FIRST ----------------------------------
    if advanced_sort:
        sheet_ratio = sheet_w / sheet_h if sheet_h else 1.0

        def ordering_key(p):
            bb = p["orig_bbox"]
            w = bb[2] - bb[0]
            h = bb[3] - bb[1] if (bb[3] - bb[1]) else 1e-6
            ratio = w / h
            area = calculate_bbox_area(bb)
            # Negative area for descending; small random jitter to break ties
            jitter = random.random() * 0.01
            return (-area, abs(ratio - sheet_ratio), jitter)

        parts.sort(key=ordering_key)
    else:
        # Still ignore upload order: shuffle first, then sort by area only
        parts.sort(
            key=lambda p: calculate_bbox_area(p["orig_bbox"]),
            reverse=True,
        )

    # ---- ROTATION CANDIDATES --------------------------------------------------
    if free_rotation:
        rotation_angles = list(
            range(0, 360, max(1, 360 // FREE_ROTATION_SAMPLES))
        )
    else:
        rotation_angles = rotations

    # Optional pruning: keep only angles whose bbox area is close to best
    if rotation_prune_tolerance > 0:
        for p in parts:
            origin = p.get("centroid", p.get("origin", (0, 0)))
            candidates = []
            min_area = float("inf")
            for ang in rotation_angles:
                rot_bbox = bbox_from_entities(
                    [transform_entity(e, origin, ang, 0, 0) for e in p["entities"]]
                )
                area = calculate_bbox_area(rot_bbox)
                candidates.append((ang, rot_bbox, area))
                if area < min_area:
                    min_area = area
            allowed = [
                ang
                for ang, _rb, area in candidates
                if area <= min_area * (1 + rotation_prune_tolerance)
            ] or rotation_angles
            p["rotation_candidates"] = allowed
    else:
        for p in parts:
            p["rotation_candidates"] = rotation_angles

    # ---- INITIAL BOTTOM‑LEFT NESTING -----------------------------------------
    sheets: list[list[dict]] = []

    while parts:
        part = parts.pop(0)
        placed = False

        # Try existing sheets first
        for sheet in sheets:
            placed_bboxes = [x["placed_bbox"] for x in sheet]
            rotation_list = part.get("rotation_candidates", rotation_angles)
            placement = bottom_left_place_with_rotation(
                part,
                placed_bboxes,
                sheet_w,
                sheet_h,
                spacing,
                rotation_list,
                aggressive=aggressive_packing,
            )
            if placement:
                rotation, candidate, anchor, rot_bbox = placement
                inst = part.copy()
                inst.update(
                    {
                        "placed_bbox": candidate,
                        "rotation": rotation,
                        "anchor_X": anchor[0],
                        "anchor_Y": anchor[1],
                        "rotation_origin": part["centroid"],
                        "rotated_min": (rot_bbox[0], rot_bbox[1]),
                    }
                )
                sheet.append(inst)
                placed = True
                break

        # If not placed, open a new sheet
        if not placed:
            rotation_list = part.get("rotation_candidates", rotation_angles)
            placement = bottom_left_place_with_rotation(
                part,
                [],
                sheet_w,
                sheet_h,
                spacing,
                rotation_list,
                aggressive=aggressive_packing,
            )
            if placement:
                rotation, candidate, anchor, rot_bbox = placement
                inst = part.copy()
                inst.update(
                    {
                        "placed_bbox": candidate,
                        "rotation": rotation,
                        "anchor_X": anchor[0],
                        "anchor_Y": anchor[1],
                        "rotation_origin": part["centroid"],
                        "rotated_min": (rot_bbox[0], rot_bbox[1]),
                    }
                )
                sheets.append([inst])
            else:
                # Part cannot fit even on an empty sheet with spacing
                debug(f"Part {part.get('label', '')} does not fit on sheet.")
                # We keep going; this part simply isn't nestable.

    # ---- COMPACTION: SLIDE LEFT/UP -------------------------------------------
    if enable_compaction:

        def _compact_sheet(sheet: list[dict]):
            # Coarse snap based on neighbours
            for _ in range(3 if aggressive_packing else 2):
                for part in sheet:
                    bbox = part["placed_bbox"]
                    w = bbox[2] - bbox[0]
                    h = bbox[3] - bbox[1]
                    others = [p["placed_bbox"] for p in sheet if p is not part]

                    # Left limit
                    left_limit = spacing
                    for ob in others:
                        if (
                            ob[2] <= bbox[0]
                            and not (ob[3] <= bbox[1] or ob[1] >= bbox[3])
                        ):
                            left_limit = max(left_limit, ob[2] + spacing)
                    new_x = max(left_limit, spacing)

                    # Top limit (CAD coords)
                    top_limit = spacing
                    for ob in others:
                        if (
                            ob[3] <= bbox[1]
                            and not (ob[2] <= bbox[0] or ob[0] >= bbox[2])
                        ):
                            top_limit = max(top_limit, ob[3] + spacing)
                    new_y = max(top_limit, spacing)

                    if new_x < bbox[0] or new_y < bbox[1]:
                        part["placed_bbox"] = (new_x, new_y, new_x + w, new_y + h)
                        part["anchor_X"] = new_x
                        part["anchor_Y"] = new_y

            # Fine sliding in small steps
            step = max(1.0, spacing)
            for part in sheet:
                moved = True
                while moved:
                    moved = False
                    bx = part["placed_bbox"]
                    w = bx[2] - bx[0]
                    h = bx[3] - bx[1]
                    others = [p["placed_bbox"] for p in sheet if p is not part]

                    # Slide horizontally
                    target_x = bx[0]
                    while target_x - step >= spacing:
                        cand = (target_x - step, bx[1], target_x - step + w, bx[3])
                        if all(not boxes_intersect(cand, o) for o in others):
                            target_x -= step
                        else:
                            break

                    # Slide vertically
                    target_y = bx[1]
                    while target_y - step >= spacing:
                        cand = (bx[0], target_y - step, bx[2], target_y - step + h)
                        if all(not boxes_intersect(cand, o) for o in others):
                            target_y -= step
                        else:
                            break

                    if target_x < bx[0] or target_y < bx[1]:
                        part["placed_bbox"] = (
                            target_x,
                            target_y,
                            target_x + w,
                            target_y + h,
                        )
                        part["anchor_X"] = target_x
                        part["anchor_Y"] = target_y
                        moved = True

        for sh in sheets:
            _compact_sheet(sh)
    else:
        _compact_sheet = lambda _sheet: None  # type: ignore

    # ---- STAGE 1: MOVE PARTS FROM LATER SHEETS TO EARLIER --------------------
    if aggressive_packing and len(sheets) > 1:
        for s_idx in range(1, len(sheets)):
            current = sheets[s_idx]
            remaining = []
            for part in current:
                relocated = False
                for earlier in sheets[:s_idx]:
                    rotation_list = part.get("rotation_candidates", rotation_angles)
                    placement = bottom_left_place_with_rotation(
                        part,
                        [p["placed_bbox"] for p in earlier],
                        sheet_w,
                        sheet_h,
                        spacing,
                        rotation_list,
                        aggressive=True,
                    )
                    if placement:
                        rotation, candidate, anchor, rot_bbox = placement
                        inst = part.copy()
                        inst.update(
                            {
                                "placed_bbox": candidate,
                                "rotation": rotation,
                                "anchor_X": anchor[0],
                                "anchor_Y": anchor[1],
                                "rotation_origin": part.get(
                                    "centroid", part.get("origin", (0, 0))
                                ),
                                "rotated_min": (rot_bbox[0], rot_bbox[1]),
                            }
                        )
                        earlier.append(inst)
                        relocated = True
                        break
                if not relocated:
                    remaining.append(part)
            sheets[s_idx] = remaining

        # Remove empty sheets and compact again
        sheets = [sh for sh in sheets if sh]
        if enable_compaction:
            for sh in sheets:
                _compact_sheet(sh)

    # ---- STAGE 2: GLOBAL MAXIMAL‑RECTANGLES REPACK ---------------------------
    if aggressive_packing and len(sheets) > 1:
        total_original_parts = sum(len(sh) for sh in sheets)
        all_parts_flat = [p for sh in sheets for p in sh]

        def _repack_with_free_rects(all_parts, sheet_w, sheet_h, spacing):
            """
            Rectangle-based repack. If *any* part cannot be placed,
            we return None so the caller can keep the original layout.
            """
            def part_bbox_at_rotation(part, ang):
                origin = part.get("centroid", part.get("origin", (0, 0)))
                rot_bbox = bbox_from_entities(
                    [
                        transform_entity(e, origin, ang, 0, 0)
                        for e in part.get("entities", [])
                    ]
                )
                w = rot_bbox[2] - rot_bbox[0]
                h = rot_bbox[3] - rot_bbox[1]
                return rot_bbox, w, h

            parts_local = [p.copy() for p in all_parts]
            for p in parts_local:
                bb = p.get("orig_bbox") or bbox_from_entities(
                    p.get("entities", [])
                )
                p["_area"] = calculate_bbox_area(bb)
            parts_local.sort(key=lambda x: x.get("_area", 0), reverse=True)

            def _init_sheet():
                return {
                    "parts": [],
                    "free_rects": [
                        (
                            spacing,
                            spacing,
                            sheet_w - 2 * spacing,
                            sheet_h - 2 * spacing,
                        )
                    ],
                }

            def _split_free_rect(free_rects, index, anchor_x, anchor_y, used_w, used_h):
                fr_x, fr_y, fr_w, fr_h = free_rects[index]
                del free_rects[index]

                rel_x = anchor_x - fr_x
                rel_y = anchor_y - fr_y

                placed_right = rel_x + used_w
                placed_bottom = rel_y + used_h

                # Left slice
                if rel_x - spacing > 0:
                    left_rect = (fr_x, fr_y, rel_x - spacing, fr_h)
                    lw = left_rect[2] - left_rect[0]
                    lh = left_rect[3] - left_rect[1]
                    if lw > 2 * spacing and lh > 2 * spacing:
                        free_rects.append(left_rect)

                # Right slice
                if fr_w - placed_right - spacing > 0:
                    right_rect = (
                        fr_x + placed_right + spacing,
                        fr_y,
                        fr_w - placed_right - spacing,
                        fr_h,
                    )
                    rw = right_rect[2] - right_rect[0]
                    rh = right_rect[3] - right_rect[1]
                    if rw > 2 * spacing and rh > 2 * spacing:
                        free_rects.append(right_rect)

                # Top slice
                if rel_y - spacing > 0:
                    top_rect = (anchor_x, fr_y, used_w, rel_y - spacing)
                    tw = top_rect[2] - top_rect[0]
                    th = top_rect[3] - top_rect[1]
                    if tw > 2 * spacing and th > 2 * spacing:
                        free_rects.append(top_rect)

                # Bottom slice
                if fr_h - placed_bottom - spacing > 0:
                    bottom_rect = (
                        anchor_x,
                        fr_y + placed_bottom + spacing,
                        used_w,
                        fr_h - placed_bottom - spacing,
                    )
                    bw = bottom_rect[2] - bottom_rect[0]
                    bh = bottom_rect[3] - bottom_rect[1]
                    if bw > 2 * spacing and bh > 2 * spacing:
                        free_rects.append(bottom_rect)

            def _prune_and_merge(free_rects):
                # Remove contained rectangles
                pruned = []
                for i, r in enumerate(free_rects):
                    rx, ry, rw, rh = r
                    contained = False
                    for j, r2 in enumerate(free_rects):
                        if i == j:
                            continue
                        rx2, ry2, rw2, rh2 = r2
                        if (
                            rx >= rx2
                            and ry >= ry2
                            and rx + rw <= rx2 + rw2
                            and ry + rh <= ry2 + rh2
                        ):
                            contained = True
                            break
                    if not contained:
                        pruned.append(r)

                # Merge touching rectangles
                merged = True
                while merged:
                    merged = False
                    out = []
                    used = [False] * len(pruned)
                    for i, a in enumerate(pruned):
                        if used[i]:
                            continue
                        ax, ay, aw, ah = a
                        merged_any = False
                        for j, b in enumerate(pruned):
                            if i == j or used[j]:
                                continue
                            bx, by, bw, bh = b
                            # Horizontal merge
                            if ay == by and ah == bh and ax + aw + spacing == bx:
                                a = (ax, ay, aw + spacing + bw, ah)
                                used[j] = True
                                merged_any = True
                            # Vertical merge
                            elif ax == bx and aw == bw and ay + ah + spacing == by:
                                a = (ax, ay, aw, ah + spacing + bh)
                                used[j] = True
                                merged_any = True
                        used[i] = True
                        out.append(a)
                        if merged_any:
                            merged = True
                    pruned = out
                return pruned

            new_sheets = []
            unplaced = []

            for part in parts_local:
                rot_list = part.get("rotation_candidates", rotations)
                best = None

                # Try to place into existing sheets
                for si, sheet in enumerate(new_sheets):
                    for fri, fr in enumerate(sheet["free_rects"]):
                        fr_x, fr_y, fr_w, fr_h = fr
                        for ang in rot_list:
                            rot_bbox, pw, ph = part_bbox_at_rotation(part, ang)
                            if pw + spacing <= fr_w and ph + spacing <= fr_h:
                                waste = (fr_w * fr_h) - (pw * ph)
                                short_fit = min(fr_w - pw, fr_h - ph)
                                long_fit = max(fr_w - pw, fr_h - ph)
                                key = (waste, short_fit, long_fit)
                                if (best is None) or (
                                    key < (best[5], best[6], best[7])
                                ):
                                    best = (
                                        si,
                                        fri,
                                        ang,
                                        pw,
                                        ph,
                                        waste,
                                        short_fit,
                                        long_fit,
                                        rot_bbox,
                                    )

                if best is None:
                    # Open new sheet
                    sheet = _init_sheet()
                    new_sheets.append(sheet)
                    chosen_ang = None
                    chosen_bbox = None
                    for ang in rot_list:
                        rot_bbox, pw, ph = part_bbox_at_rotation(part, ang)
                        if (
                            pw + spacing <= sheet_w - 2 * spacing
                            and ph + spacing <= sheet_h - 2 * spacing
                        ):
                            chosen_ang = ang
                            chosen_bbox = (rot_bbox, pw, ph)
                            break
                    if chosen_ang is None:
                        unplaced.append(part)
                        continue  # cannot fit this part at all

                    rot_bbox, pw, ph = chosen_bbox
                    anchor_x = spacing
                    anchor_y = spacing
                    placed_bbox = (
                        anchor_x,
                        anchor_y,
                        anchor_x + pw,
                        anchor_y + ph,
                    )
                    inst = part.copy()
                    inst.update(
                        {
                            "placed_bbox": placed_bbox,
                            "rotation": chosen_ang,
                            "anchor_X": anchor_x,
                            "anchor_Y": anchor_y,
                            "rotation_origin": part.get(
                                "centroid", part.get("origin", (0, 0))
                            ),
                            "rotated_min": (rot_bbox[0], rot_bbox[1]),
                        }
                    )
                    sheet["parts"].append(inst)
                    _split_free_rect(sheet["free_rects"], 0, anchor_x, anchor_y, pw, ph)
                    sheet["free_rects"] = _prune_and_merge(sheet["free_rects"])
                else:
                    si, fri, ang, pw, ph, waste, short_fit, long_fit, rot_bbox = best
                    sheet = new_sheets[si]
                    fr_x, fr_y, fr_w, fr_h = sheet["free_rects"][fri]

                    corner_candidates = []
                    possible_corners = [
                        (fr_x, fr_y),
                        (fr_x + max(0, fr_w - pw - spacing), fr_y),
                        (fr_x, fr_y + max(0, fr_h - ph - spacing)),
                    ]
                    for ax, ay in possible_corners:
                        if (
                            ax + pw + spacing <= fr_x + fr_w
                            and ay + ph + spacing <= fr_y + fr_h
                        ):
                            frag_score = abs((fr_w * fr_h) - (pw * ph))
                            corner_candidates.append((frag_score, ax, ay))

                    if corner_candidates:
                        corner_candidates.sort(key=lambda x: x[0])
                        _, anchor_x, anchor_y = corner_candidates[0]
                    else:
                        anchor_x = fr_x
                        anchor_y = fr_y

                    placed_bbox = (
                        anchor_x,
                        anchor_y,
                        anchor_x + pw,
                        anchor_y + ph,
                    )
                    inst = part.copy()
                    inst.update(
                        {
                            "placed_bbox": placed_bbox,
                            "rotation": ang,
                            "anchor_X": anchor_x,
                            "anchor_Y": anchor_y,
                            "rotation_origin": part.get(
                                "centroid", part.get("origin", (0, 0))
                            ),
                            "rotated_min": (rot_bbox[0], rot_bbox[1]),
                        }
                    )
                    sheet["parts"].append(inst)
                    _split_free_rect(
                        sheet["free_rects"], fri, anchor_x, anchor_y, pw, ph
                    )
                    sheet["free_rects"] = _prune_and_merge(sheet["free_rects"])

            if unplaced:
                # Repack failed to place some parts – signal failure
                debug(f"Repack could not place {len(unplaced)} parts; keeping original layout.")
                return None

            return [s["parts"] for s in new_sheets if s["parts"]]

        repacked_sheets = _repack_with_free_rects(
            all_parts_flat, sheet_w, sheet_h, spacing
        )

        if repacked_sheets is not None:
            total_repacked = sum(len(sh) for sh in repacked_sheets)
            if (
                len(repacked_sheets) <= len(sheets)
                and total_repacked == total_original_parts
            ):
                # Accept improved layout
                sheets = repacked_sheets
                if enable_compaction:
                    for sh in sheets:
                        _compact_sheet(sh)

    return sheets


def write_sheets_entities_to_dxf(sheets, base_output, sheet_w, sheet_h):
    for idx, sheet in enumerate(sheets, start=1):
        doc = getattr(ezdxf, "new")("R2010")
        msp = doc.modelspace()
        msp.add_lwpolyline(
            [(0, 0), (sheet_w, 0), (sheet_w, sheet_h), (0, sheet_h)],
            close=True,
            dxfattribs={"layer": "SHEET"},
        )
        for part in sheet:
            ang = part.get("rotation", 0)
            anchor_x = part.get("anchor_X", 0)
            anchor_y = part.get("anchor_Y", 0)
            origin = part.get("rotation_origin", (0, 0))
            rot_min = part.get("rotated_min", (0, 0))
            tx = anchor_x - rot_min[0]
            ty = anchor_y - rot_min[1]

            if part.get("outer"):
                o = transform_entity(part["outer"], origin, ang, tx, ty)
                if o and o["type"] == "LWPOLYLINE":
                    pts = [(p[0], p[1]) for p in o["points"]]
                    msp.add_lwpolyline(pts, close=o.get("closed", False))

            for (ent, offset) in part.get("relative_offsets", []):
                rot_off = rotate_point_around((0, 0), offset, ang)
                tx2 = tx + rot_off[0]
                ty2 = ty + rot_off[1]
                ent2 = transform_entity(ent, origin, ang, tx2, ty2)
                if not ent2:
                    continue
                if ent2["type"] == "LWPOLYLINE":
                    msp.add_lwpolyline(
                        ent2["points"], close=ent2.get("closed", False)
                    )
                elif ent2["type"] == "CIRCLE":
                    msp.add_circle(ent2["center"], ent2["radius"])
                elif ent2["type"] == "ARC":
                    msp.add_arc(
                        ent2["center"],
                        ent2["radius"],
                        ent2["start_angle"],
                        ent2["end_angle"],
                    )
                elif ent2["type"] == "SPLINE":
                    msp.add_lwpolyline(ent2["points"], close=False)

            for ent in part.get("entities", []):
                if ent in [part.get("outer")] + [
                    e for e, _ in part.get("relative_offsets", [])
                ]:
                    continue
                t = ent["type"]
                e2 = transform_entity(ent, origin, ang, tx, ty)
                if not e2:
                    continue
                if t == "LWPOLYLINE":
                    msp.add_lwpolyline(e2["points"], close=e2.get("closed", False))
                elif t == "CIRCLE":
                    msp.add_circle(e2["center"], e2["radius"])
                elif t == "LINE":
                    msp.add_line(e2["start"], e2["end"])
                elif t == "ARC":
                    msp.add_arc(
                        e2["center"],
                        e2["radius"],
                        e2["start_angle"],
                        e2["end_angle"],
                    )
                elif t == "SPLINE":
                    msp.add_lwpolyline(e2["points"], close=False)
        out_path = f"{base_output}_sheet{idx:03d}.dxf"
        try:
            doc.saveas(out_path)
            print("Saved:", out_path)
        except Exception as ex:
            print("Error saving:", ex)


# -----------------------
# Streamlit UI
# -----------------------
logo_candidates = [
    r"D:\\Downloads\\.png",
    r"D:\\Downloads\\ei8-logo.png",
    "d:/Downloads/ei8-logo.png",
]
logo_path = next((p for p in logo_candidates if os.path.exists(p)), None)


def _logo_data_uri(path: str | None) -> str | None:
    if not path:
        return None
    try:
        with open(path, "rb") as f:
            data = f.read()
        return "data:image/png;base64," + base64.b64encode(data).decode("ascii")
    except Exception:
        return None


logo_src = _logo_data_uri(logo_path)
page_icon = None
try:
    page_icon = Image.open(logo_path) if logo_path else None
except Exception:
    page_icon = None

st.set_page_config(
    page_title="Ei8 DXF Nesting",
    page_icon=(page_icon or "📐"),
    layout="wide",
)

accent_color = "#0FF0A1"
dark_mode = True
surface_bg = "#101827" if dark_mode else "#ffffff"
app_bg = "#0d1117" if dark_mode else "#f4f7fb"
text_color = "#f1f5f9" if dark_mode else "#0f172a"
subtle_text = "#94a3b8" if dark_mode else "#64748b"
border_color = "#1e293b" if dark_mode else "#e2e8f0"
grad_from = accent_color
grad_to = accent_color + "80" if len(accent_color) == 7 else accent_color

st.markdown(
    f"""
    <style>
    .stApp {{ background: {app_bg}; }}
    .title-row {{ display:flex; align-items:center; gap:12px; margin-bottom:6px; }}
    .logo-img {{ height:36px; width:auto; border-radius:6px; }}
    .panel {{ background: {surface_bg}; border:1px solid {border_color};
              border-radius: 14px; padding: 16px 18px 14px;
              box-shadow: 0 4px 12px rgba(0,0,0,0.08); }}
    .title {{ font-size: 1.55rem; font-weight:700;
              background: linear-gradient(90deg,{grad_from},{grad_to});
              -webkit-background-clip: text; color:transparent; margin-bottom:10px; }}
    .muted {{ color:{subtle_text}; font-size:13px; }}
    .pill {{ display:inline-block; padding:2px 10px; border-radius:999px;
             background:{accent_color}20; color:{accent_color}; font-size:11px;
             margin:3px 6px 0 0; letter-spacing:.5px; }}
    .file-card {{ border:1px solid {border_color}; background:{surface_bg};
                 padding:12px 16px 14px; border-radius:12px; margin-bottom:12px;
                 position:relative; }}
    .file-card:hover {{ box-shadow:0 4px 10px rgba(0,0,0,0.12);
                        border-color:{accent_color}; }}
    .file-title {{ font-weight:600; color:{text_color}; font-size:15px; }}
    .file-meta {{ color:{subtle_text}; font-size:12px; margin-left:6px; }}
    .nav-area .stButton > button {{ min-width:0 !important; width:100% !important; }}
    .metric-row div[data-testid="stMetricValue"] {{ color:{accent_color}; }}

    /* Parts table styles */
    .parts-table {{ width:100%; border-collapse: collapse; }}
    .parts-table th, .parts-table td {{ border-bottom:1px solid {border_color}; padding:8px 10px; vertical-align:top; }}
    .parts-table th {{ color:{subtle_text}; font-weight:600; text-align:left; font-size:13px; }}
    .parts-thumb {{ width:140px; height:80px; object-fit:contain; border:1px solid {border_color}; border-radius:6px; background:#fff; }}
    .parts-filename {{ font-weight:600; color:{text_color}; font-size:14px; }}
    .qty-wrap .stButton>button {{ padding:4px 10px; }}
    .qty-cell .stButton > button {{
        background: transparent !important;
        color: {text_color} !important;
        border: 1px solid {border_color} !important;
        padding: 2px 8px !important;
        min-width: 32px !important;
        height: 28px !important;
        line-height: 1 !important;
        border-radius: 6px !important;
        box-shadow: none !important;
    }}
    .qty-cell .stButton > button:hover {{ border-color: {accent_color} !important; }}
    .rm-cell .stButton > button {{
        background:{accent_color} !important;
        color:#0b1220 !important;
        border:none !important;
        white-space: nowrap !important;
        font-weight:600 !important;
        padding: 6px 16px !important; /* reduced for vertical alignment */
        border-radius:10px !important;
        min-width: 110px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        position: relative !important;
        top: -2px !important; /* shift upward */
    }}

    @media (max-width: 1100px) {{
        .stButton > button {{ padding:7px 18px; min-width:100px; }}
    }}
    @media (max-width: 900px) {{
        .panel {{ padding:14px 14px 12px; }}
        .file-card {{ padding:10px 12px; }}
    }}
    @media (max-width: 700px) {{
        .title {{ font-size:1.3rem; }}
        .stButton > button {{ padding:6px 14px; font-size:13px; }}
    }}
    @media (max-width: 480px) {{
        .title {{ font-size:1.15rem; }}
        .stButton > button {{ padding:6px 12px; font-size:12px; min-width:unset; }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def safe_read_dxf_from_upload(uploaded_file):
    uploaded_file.seek(0)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name
    doc = getattr(ezdxf, "readfile")(tmp_path)
    return doc, tmp_path


def create_sheet_preview_image(
    sheet,
    sheet_w,
    sheet_h,
    canvas_px=900,
    margin=20,
    draw_bbox=True,
    draw_outline=True,
):
    if sheet_w <= 0 or sheet_h <= 0:
        return None
    aspect = sheet_w / sheet_h
    if aspect >= 1:
        w_px = canvas_px
        h_px = int(canvas_px / aspect)
    else:
        h_px = canvas_px
        w_px = int(canvas_px * aspect)
    w_px = max(320, w_px)
    h_px = max(240, h_px)
    # Use asymmetric vertical margins to reduce top whitespace
    margin_top = int(margin * 0.35)
    margin_bottom = margin * 2 - margin_top
    margin_side = margin
    img_w = w_px + margin_side * 2
    img_h = h_px + margin_top + margin_bottom
    img = Image.new("RGB", (img_w, img_h), (250, 251, 253))
    draw = ImageDraw.Draw(img)
    sx = w_px / sheet_w
    sy = h_px / sheet_h
    sheet_x1, sheet_y1 = margin_side, margin_top
    sheet_x2, sheet_y2 = margin_side + w_px, margin_top + h_px
    draw.rectangle(
        [sheet_x1, sheet_y1, sheet_x2, sheet_y2],
        outline=(190, 200, 210),
        width=2,
        fill=(255, 255, 255),
    )
    # Pastel color palette (light tones); extend with random variation for more sheets
    base_palette = [
        (197, 225, 165),  # soft green
        (187, 212, 237),  # soft blue
        (232, 205, 192),  # soft clay
        (203, 191, 235),  # soft violet
        (180, 224, 223),  # soft teal
        (246, 222, 180),  # soft apricot
    ]
    # Slightly shuffle palette for visual variety between renders
    palette = base_palette[:]
    random.shuffle(palette)

    for idx, part in enumerate(sheet):
        bbox = part.get("placed_bbox", (0, 0, 0, 0))
        x1, y1, x2, y2 = bbox
        # Invert Y so higher CAD Y appears lower on screen (downward flip)
        px1 = margin_side + int(x1 * sx)
        py1 = margin_top + int((sheet_h - y2) * sy)
        px2 = margin_side + int(x2 * sx)
        py2 = margin_top + int((sheet_h - y1) * sy)

        if draw_bbox:
            color = palette[idx % len(palette)]
            draw.rectangle(
                [px1, py1, px2, py2],
                outline=(255, 255, 255),
                width=2,
                fill=color,
            )
        # Removed part name labeling per user request

        outer = part.get("outer")
        if outer and isinstance(outer, dict):
            pts = outer.get("points") or []
            if pts:
                origin = part.get("rotation_origin", (0, 0))
                ang = part.get("rotation", 0)
                rot_min = part.get("rotated_min", (0, 0))
                tx = part.get("anchor_X", 0) - rot_min[0]
                ty = part.get("anchor_Y", 0) - rot_min[1]
                pts_px = []
                for p in pts:
                    x_tr, y_tr = transform_point(p, origin, ang, tx, ty)
                    px = margin_side + int(x_tr * sx)
                    py = margin_top + int((sheet_h - y_tr) * sy)
                    pts_px.append((px, py))
                if draw_outline and len(pts_px) >= 2:
                    draw.line(
                        pts_px + [pts_px[0]], fill=(80, 110, 140), width=1
                    )

        if draw_outline:
            origin = part.get("rotation_origin", (0, 0))
            ang = part.get("rotation", 0)
            rot_min = part.get("rotated_min", (0, 0))
            tx = part.get("anchor_X", 0) - rot_min[0]
            ty = part.get("anchor_Y", 0) - rot_min[1]
            for ent in part.get("entities", []):
                if outer and ent is outer:
                    continue
                t = ent.get("type")
                if t == "LWPOLYLINE" and ent.get("points"):
                    pts_px = []
                    for p in ent["points"]:
                        x_tr, y_tr = transform_point(p, origin, ang, tx, ty)
                        px = margin_side + int(x_tr * sx)
                        py = margin_top + int((sheet_h - y_tr) * sy)
                        pts_px.append((px, py))
                    if len(pts_px) >= 2:
                        draw.line(
                            pts_px + [pts_px[0]], fill=(110, 130, 160), width=1
                        )
                elif t == "LINE":
                    s = ent.get("start")
                    e = ent.get("end")
                    if s and e:
                        s_tr = transform_point(s, origin, ang, tx, ty)
                        e_tr = transform_point(e, origin, ang, tx, ty)
                        draw.line(
                            [
                                (
                                    margin_side + int(s_tr[0] * sx),
                                    margin_top + int((sheet_h - s_tr[1]) * sy),
                                ),
                                (
                                    margin_side + int(e_tr[0] * sx),
                                    margin_top + int((sheet_h - e_tr[1]) * sy),
                                ),
                            ],
                            fill=(110, 130, 160),
                            width=1,
                        )
                elif t == "CIRCLE":
                    c = transform_point(ent["center"], origin, ang, tx, ty)
                    r = ent.get("radius", 0)
                    cx = margin_side + int(c[0] * sx)
                    cy = margin_top + int((sheet_h - c[1]) * sy)
                    rr = int(r * (sx + sy) / 2)
                    draw.ellipse(
                        [cx - rr, cy - rr, cx + rr, cy + rr],
                        outline=(110, 130, 160),
                        width=1,
                    )
    return img


@st.cache_data(show_spinner=False)
def create_file_preview_thumbnail(meta: dict, width: int = 140, height: int = 80) -> Image.Image:
    parts = meta.get("parts", []) or []
    img = Image.new("RGB", (width, height), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    if not parts:
        # empty placeholder
        draw.rectangle([1, 1, width - 2, height - 2], outline=(200, 210, 220))
        return img

    # Use first part for quick preview
    p = parts[0]
    ents = p.get("entities", []) or []
    if not ents and p.get("outer"):
        ents = [p["outer"]]

    bbox = p.get("orig_bbox") or bbox_from_entities(ents) or (0, 0, 0, 0)
    x1, y1, x2, y2 = bbox
    bw = max(1e-6, (x2 - x1))
    bh = max(1e-6, (y2 - y1))
    pad_x, pad_y = 6, 6
    sx = max(1e-6, (width - 2 * pad_x) / bw)
    sy = max(1e-6, (height - 2 * pad_y) / bh)
    s = min(sx, sy)

    drew = False
    outer = p.get("outer") if isinstance(p.get("outer"), dict) else None
    if outer and outer.get("type") == "LWPOLYLINE" and outer.get("points"):
        pts = []
        for pt in outer["points"]:
            px = pad_x + int((pt[0] - x1) * s)
            py = height - pad_y - int((pt[1] - y1) * s)
            pts.append((px, py))
        if len(pts) >= 2:
            draw.line(pts + [pts[0]], fill=(20, 80, 160), width=2)
            drew = True

    if not drew:
        for ent in ents:
            t = ent.get("type")
            if t == "LWPOLYLINE" and ent.get("points"):
                pts = []
                for pt in ent["points"]:
                    px = pad_x + int((pt[0] - x1) * s)
                    py = height - pad_y - int((pt[1] - y1) * s)
                    pts.append((px, py))
                if len(pts) >= 2:
                    draw.line(pts + [pts[0]], fill=(50, 120, 200), width=1)
                    drew = True
            elif t == "LINE":
                spt = ent.get("start")
                ept = ent.get("end")
                if spt and ept:
                    xA = pad_x + int((spt[0] - x1) * s)
                    yA = height - pad_y - int((spt[1] - y1) * s)
                    xB = pad_x + int((ept[0] - x1) * s)
                    yB = height - pad_y - int((ept[1] - y1) * s)
                    draw.line([(xA, yA), (xB, yB)], fill=(60, 130, 210), width=1)
                    drew = True

    if not drew:
     
        draw.rectangle([pad_x, pad_y, width - pad_x, height - pad_y], outline=(200, 210, 220))
    return img

if "uploaded_meta" not in st.session_state:
    st.session_state.uploaded_meta = []
if "nested_sheets" not in st.session_state:
    st.session_state.nested_sheets = None
if "config" not in st.session_state:
    st.session_state.config = {}
if "uploader_snapshot" not in st.session_state:
    st.session_state.uploader_snapshot = None
if "ignore_upload_names" not in st.session_state:
    st.session_state.ignore_upload_names = set()
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = f"uploader_{uuid.uuid4()}"
if "part_summary" not in st.session_state:
    st.session_state.part_summary = []
if "stage" not in st.session_state:
    st.session_state.stage = "SET"
if "project_name" not in st.session_state and st.session_state.get("project_id"):
    meta = _load_project_meta(st.session_state.project_id)
    raw = (meta or {}).get("name")
    st.session_state.project_name = _normalized_project_name(raw, st.session_state.project_id)
if "project_created_at" not in st.session_state and st.session_state.get("project_id"):
    meta = _load_project_meta(st.session_state.project_id)
    st.session_state.project_created_at = (meta or {}).get("created_at", datetime.datetime.now().isoformat())
if "export_sheet_number" not in st.session_state:
    st.session_state.export_sheet_number = 1
if "sheet_nav_number" not in st.session_state:
    st.session_state.sheet_nav_number = 1
if "config_inputs" not in st.session_state:
    st.session_state.config_inputs = {
        "sheet_w": DEFAULT_SHEET_W,
        "sheet_h": DEFAULT_SHEET_H,
        "spacing": DEFAULT_SPACING,
        "kerf": DEFAULT_KERF,
        "rotation_mode": "Free Rotation (Optimized)",
        "rotation_step": DEFAULT_ROT_STEP,
        "preview_canvas_px": 900,
        "show_bbox": True,
        "show_outline": True,
        "advanced_sort": True,
        "enable_compaction": True,
        "aggressive_packing": True,
    }

_init_project_if_needed()

with st.sidebar:
    st.markdown("### Projects")
    _ensure_projects_root()
    existing = [d for d in os.listdir(PROJECTS_ROOT) if os.path.isdir(os.path.join(PROJECTS_ROOT, d))]
    existing.sort(reverse=True)
    # Build list with metadata
    project_items = []
    for pid in existing:
        meta = _load_project_meta(pid) or {}
        created_raw = meta.get("created_at")
        if not created_raw:
            continue
        name = _normalized_project_name(meta.get("name"), pid)
        try:
            dt_obj = datetime.datetime.fromisoformat(created_raw)
            created_h = dt_obj.strftime("%b %d, %Y — %H:%M")
        except Exception:
            # Skip entries we cannot confidently parse for date/time
            continue
        project_items.append({"pid": pid, "name": name, "created_h": created_h})
    if project_items:
        for item in project_items:
            label = item["name"]
            if item["created_h"]:
                label = f"{label} · {item['created_h']}"
            item["label"] = label

        project_selector_values = ["__current__"] + [item["pid"] for item in project_items]

        def _project_option_label(value):
            if value == "__current__":
                return "(current)"
            match = next((itm for itm in project_items if itm["pid"] == value), None)
            return match["label"] if match else DEFAULT_PROJECT_DISPLAY_NAME

        sel = st.selectbox(
            "Switch Project",
            project_selector_values,
            index=0,
            key="proj_select",
            format_func=_project_option_label,
        )
        prev_sel = st.session_state.get("_prev_proj_sel", "__current__")
        if sel != prev_sel:
            st.session_state._prev_proj_sel = sel
            if sel != "__current__":
                if _load_project(sel):
                    label = _project_option_label(sel)
                    st.success(f"Loaded '{label}'")
                    st.toast(f"Project '{label}' loaded", icon="📁")
                    st.rerun()
        with st.expander("Project Details", expanded=False):
            for item in project_items:
                cur = item["pid"] == st.session_state.get("project_id")
                marker = " (current)" if cur else ""
                cols = st.columns([4,1])
                with cols[0]:
                    st.markdown(f"**{item['name']}**{marker}<br><span style='font-size:11px;color:#8892a2'>{item['created_h']}</span>", unsafe_allow_html=True)
                with cols[1]:
                    btn_label = "🗑" if not cur else "🗑"
                    if st.button(btn_label, key=f"del_{item['pid']}"):
                        if _delete_project(item['pid']):
                            # If current deleted, create replacement
                            if cur:
                                _create_new_project(DEFAULT_PROJECT_DISPLAY_NAME)
                            st.toast("Project deleted", icon="❌")
                            st.rerun()
                st.markdown("<hr style='margin:6px 0'>", unsafe_allow_html=True)
    if st.button("➕ New Project", key="proj_new_btn", use_container_width=True):
        _create_new_project()
        st.toast("New project created", icon="🆕")
        st.rerun()

    def _rename_current():
        new_name = st.session_state.get("proj_name_input")
        if new_name:
            trimmed = new_name.strip()
        else:
            trimmed = ""
        if trimmed and trimmed != st.session_state.get("project_name"):
            st.session_state.project_name = _normalized_project_name(trimmed, st.session_state.get("project_id"))
            _save_project_meta()
            _auto_save_project()
    active_project_name = _normalized_project_name(st.session_state.get("project_name"), st.session_state.get("project_id"))
    if st.session_state.get("project_name") != active_project_name:
        st.session_state.project_name = active_project_name
    st.text_input("Project Name", value=active_project_name, key="proj_name_input", on_change=_rename_current)
    # Show only human-friendly project name, not internal ID
    st.caption(f"Active project: {active_project_name}")
    if st.button("Force Save", key="proj_force_save"):
        _auto_save_project("manual")
        st.success("Project saved")
        st.toast("💾 Saved", icon="💾")


def go_to_stage(target: str):
    st.session_state.stage = target


stage = st.session_state.stage

header_html = (
    f"<div class='title-row'><img src='{logo_src}' class='logo-img' alt='Ei8 logo'/>"
    f"<div class='title'>DXF Nesting</div></div>"
    if logo_src
    else "<div class='title'>DXF Nesting</div>"
)
st.markdown(header_html, unsafe_allow_html=True)
st.markdown(
    "<div class='muted'>Three-stage flow: "
    "<strong>SET</strong> parts • "
    "<strong>NEST</strong> configuration • "
    "<strong>CUT</strong> results & export.</div>",
    unsafe_allow_html=True,
)

stage_container = st.container()
with stage_container:
    st.markdown("<div class='stage-tabs'>", unsafe_allow_html=True)
    tab_cols = st.columns(3)
    order = ["SET", "NEST", "CUT"]
    for i, key in enumerate(order):
        with tab_cols[i]:
            st.button(
                key,
                key=f"tab_{key}",
                use_container_width=True,
                on_click=go_to_stage,
                args=(key,),
            )
    st.markdown("</div>", unsafe_allow_html=True)
    active_idx = order.index(stage) + 1
    st.markdown(f"""
    <style>
    .stage-tabs > div > div > div > button {{
        background: transparent !important;
        border: 1px solid {border_color} !important;
        color: {text_color} !important;
        font-weight: 600 !important;
        box-shadow: none !important;
    }}
    .stage-tabs > div:nth-child({active_idx}) > div > div > button {{
        background: {accent_color} !important;
        color: #0b1220 !important;
        border: 1px solid {accent_color} !important;
        box-shadow: 0 0 0 1px {accent_color} inset !important;
    }}
    .stage-tabs > div > div > div > button:hover {{
        border-color: {accent_color} !important;
        color: {accent_color} !important;
    }}
    </style>
    """, unsafe_allow_html=True)


col_l, col_c, col_r = st.columns([1.0, 2.8, 0.9])

with col_l:
    if stage == "SET":
        st.markdown(
            "<div class='panel'><strong>Files • Upload</strong></div>",
            unsafe_allow_html=True,
        )
        uploaded_files = st.file_uploader(
            "Upload DXF files",
            type=["dxf"],
            accept_multiple_files=True,
            help="Drag & drop multiple DXF files. We'll parse outlines and holes automatically.",
            key=st.session_state.uploader_key,
        )

        current_names = tuple(f.name for f in uploaded_files) if uploaded_files else ()
        if st.session_state.uploader_snapshot != current_names:
            st.session_state.uploader_snapshot = current_names
            st.session_state.ignore_upload_names = set()

        if uploaded_files:
            for f in uploaded_files:
                if f.name in st.session_state.ignore_upload_names:
                    continue
                already = any(
                    m["name"] == f.name
                    and m.get("size") == getattr(f, "size", None)
                    for m in st.session_state.uploaded_meta
                )
                if already:
                    continue
                meta = {
                    "id": str(uuid.uuid4()),
                    "name": f.name,
                    "tmp_path": None,
                    "parts": [],
                    "valid": False,
                    "error": None,
                    "qty": 1,
                    "label": f.name,
                }
                try:
                    doc, tmp_path = safe_read_dxf_from_upload(f)
                    meta["tmp_path"] = tmp_path

                    msp = doc.modelspace()
                    entities = []
                    for e in msp:
                        rep = transform_entity_representation(e)
                        if rep:
                            entities.append(rep)

                    try:
                        counts = detect_entity_types(doc)
                        meta["entity_counts"] = dict(counts)
                    except Exception:
                        meta["entity_counts"] = {}
                    parts = group_entities_universal(entities, doc)
                    for p in parts:
                        calculate_relative_offsets_for_part(p)
                        ents = []
                        if p.get("outer"):
                            ents.append(p["outer"])
                        ents += (
                            p.get("holes", [])
                            + p.get("circles", [])
                            + p.get("arcs", [])
                            + p.get("splines", [])
                            + p.get("others", [])
                        )
                        p["entities"] = ents
                        p["orig_bbox"] = bbox_from_entities(ents)
                    meta["parts"] = parts
                    meta["valid"] = True
                    meta["size"] = (
                        os.path.getsize(tmp_path)
                        if tmp_path and os.path.exists(tmp_path)
                        else None
                    )
                except Exception as e:
                    meta["error"] = f"{type(e).__name__}: {str(e)}"
                    meta["valid"] = False
                st.session_state.uploaded_meta.append(meta)

        st.write("")
        if not st.session_state.uploaded_meta:
            st.info("No DXF files uploaded. Use the uploader above.")

        if st.button("Reset files", use_container_width=True):
            for m in st.session_state.uploaded_meta:
                try:
                    if m.get("tmp_path") and os.path.exists(m["tmp_path"]):
                        os.remove(m["tmp_path"])
                except Exception:
                    pass
            st.session_state.uploaded_meta = []
            st.session_state.nested_sheets = None
            st.session_state.ignore_upload_names = set()
            st.session_state.uploader_snapshot = tuple()
            st.session_state.part_summary = []
         
            st.session_state.uploader_key = f"uploader_{uuid.uuid4()}"
            st.rerun()

    else:
        st.markdown(
            "<div class='panel'><strong>Files</strong></div>",
            unsafe_allow_html=True,
        )
        if not st.session_state.uploaded_meta:
            st.info("No DXF files uploaded. Go to the SET stage to add parts.")
        else:
            for m in st.session_state.uploaded_meta:
                label = m.get("label", m["name"])
                qty = m.get("qty", 1)
                st.write(f"• **{label}** × {qty}")

with col_c:
    if stage == "SET":
        st.markdown(
            "<div class='panel'><strong>Overview</strong></div>",
            unsafe_allow_html=True,
        )
        total_files = len(st.session_state.uploaded_meta)
        total_parts = sum(
            int(m.get("qty", 1)) * len(m.get("parts", []))
            for m in st.session_state.uploaded_meta
            if m.get("valid")
        )
        st.write(f"- Uploaded files: **{total_files}**")
        st.write(f"- Requested parts (before nesting): **{total_parts}**")
        st.write(
            "- Once you're done setting up files, click **Next** or the "
            "**NEST** tab to configure sheet and nesting."
        )

        st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
        if st.session_state.uploaded_meta:
            st.write("")
            st.markdown(
                "<div class='panel'><strong>Set — Parts</strong></div>",
                unsafe_allow_html=True,
            )

            col_spec = [6.2, 2.0, 1.4, 1.4, 1.9]
            h_cols = st.columns(col_spec)
            h_cols[0].markdown("**Filename / Preview**")
            h_cols[1].markdown("**Quantity**")
            h_cols[2].markdown("**Height (mm)**")
            h_cols[3].markdown("**Width (mm)**")
            h_cols[4].markdown("**Remove**")

            def _change_qty(mid: str, delta: int):
                for it in st.session_state.uploaded_meta:
                    if it["id"] == mid:
                        it["qty"] = max(1, int(it.get("qty", 1)) + delta)
                        break

            removed_id_tbl = None

            for m in st.session_state.uploaded_meta:
                r_cols = st.columns(col_spec)
                with r_cols[0]:
                    thumb = create_file_preview_thumbnail(m, width=200, height=120)
                    sub = st.columns([1.8, 4.2])
                    with sub[0]:
                        st.image(thumb, use_container_width=False, clamp=True)
                    with sub[1]:
                        new_label = st.text_input(
                            "",
                            value=m.get("label", m["name"]),
                            key=f"tbl_label_{m['id']}",
                            label_visibility="collapsed",
                            placeholder="Enter name…",
                        )
                        st.markdown(
                            f"<div class='parts-filename' style='margin-top:4px'>{m['name']}</div>",
                            unsafe_allow_html=True,
                        )

               
                with r_cols[1]:
                    # Revert to native number input (arrows inside box) for instant single-click change
                    new_qty = st.number_input(
                        "Qty",
                        min_value=1,
                        step=1,
                        value=int(m.get("qty", 1)),
                        key=f"tbl_qty_num_{m['id']}",
                        label_visibility="collapsed",
                    )
                    # Slight upward shift for alignment with header row
                    st.markdown("<div style='position:relative;top:-4px;height:0'></div>", unsafe_allow_html=True)

              
                b_w = b_h = 0.0
                parts = m.get("parts", []) or []
                if parts:
                    p0 = parts[0]
                    bb = p0.get("orig_bbox") or bbox_from_entities(p0.get("entities", []))
                    b_w = max(0.0, (bb[2] - bb[0]))
                    b_h = max(0.0, (bb[3] - bb[1]))
                with r_cols[2]:
                    st.markdown(f"{b_h:.1f}")
                with r_cols[3]:
                    st.markdown(f"{b_w:.1f}")

                # Remove
                with r_cols[4]:
                    # Shift upward for visual alignment with other row cells
                    st.markdown("<div class='rm-cell' style='position:relative;top:-6px'>", unsafe_allow_html=True)
                    if st.button(
                        "Remove",
                        key=f"tbl_rm_{m['id']}",
                    ):
                        removed_id_tbl = m["id"]
                    st.markdown("</div>", unsafe_allow_html=True)

               
                st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

                # Persist label & quantity updates
                for item in st.session_state.uploaded_meta:
                    if item["id"] == m["id"]:
                        item["label"] = new_label
                        item["qty"] = int(new_qty)
                        break

            if removed_id_tbl:
                target = None
                for item in st.session_state.uploaded_meta:
                    if item["id"] == removed_id_tbl:
                        target = item
                        break
                if target:
                    try:
                        if target.get("tmp_path") and os.path.exists(target["tmp_path"]):
                            os.remove(target["tmp_path"])
                    except Exception:
                        pass
                    st.session_state.ignore_upload_names.add(target["name"])
                    st.session_state.uploaded_meta = [
                        it for it in st.session_state.uploaded_meta if it["id"] != removed_id_tbl
                    ]
                    st.rerun()


    elif stage == "NEST":
        cfg_inputs = st.session_state.config_inputs
        st.markdown(
            "<div class='panel'><strong>Nesting configuration</strong></div>",
            unsafe_allow_html=True,
        )

        if not st.session_state.uploaded_meta:
            st.warning("No DXF files uploaded yet. Go back to **SET** first.")
        else:
            st.write(
                "Use this section to configure **Preview**, **Sheet**, and "
                "**Nesting** parameters, then click **Start Nesting**."
            )
            st.write("")

            col_cfg_left, col_cfg_right = st.columns(2)

            # Preview controls
            with col_cfg_left:
                st.markdown("#### Preview")
                cfg_inputs["preview_canvas_px"] = st.slider(
                    "Preview Size (px)",
                    480,
                    1400,
                    int(cfg_inputs.get("preview_canvas_px", 900)),
                    step=20,
                    key="preview_px",
                )
                cfg_inputs["show_bbox"] = st.toggle(
                    "Show Bounding Boxes",
                    value=bool(cfg_inputs.get("show_bbox", True)),
                    key="show_bbox_toggle",
                )
                cfg_inputs["show_outline"] = st.toggle(
                    "Show Part Outlines",
                    value=bool(cfg_inputs.get("show_outline", True)),
                    key="show_outline_toggle",
                )

            # Sheet + nesting controls
            with col_cfg_right:
                st.markdown("#### Sheet")
                cfg_inputs["sheet_w"] = st.number_input(
                    "Width (mm)",
                    value=float(cfg_inputs.get("sheet_w", DEFAULT_SHEET_W)),
                    min_value=100.0,
                    step=10.0,
                    key="sheet_w_input",
                )
                cfg_inputs["sheet_h"] = st.number_input(
                    "Height (mm)",
                    value=float(cfg_inputs.get("sheet_h", DEFAULT_SHEET_H)),
                    min_value=100.0,
                    step=10.0,
                    key="sheet_h_input",
                )

                st.markdown("#### Nesting")
                cfg_inputs["spacing"] = st.number_input(
                    "Spacing (mm)",
                    value=float(cfg_inputs.get("spacing", DEFAULT_SPACING)),
                    min_value=0.0,
                    step=0.5,
                    key="spacing_input",
                )
                cfg_inputs["kerf"] = st.number_input(
                    "Kerf (mm)",
                    value=float(cfg_inputs.get("kerf", DEFAULT_KERF)),
                    min_value=0.0,
                    step=0.1,
                    key="kerf_input",
                )

                rotation_mode_options = [
                    "Free Rotation (Optimized)",
                    "Fixed Step Rotation",
                ]
                current_mode = cfg_inputs.get(
                    "rotation_mode", rotation_mode_options[0]
                )
                mode_index = (
                    rotation_mode_options.index(current_mode)
                    if current_mode in rotation_mode_options
                    else 0
                )
                cfg_inputs["rotation_mode"] = st.radio(
                    "Rotation Mode",
                    rotation_mode_options,
                    index=mode_index,
                    key="rotation_mode_radio",
                )
                if cfg_inputs["rotation_mode"] == "Fixed Step Rotation":
                    cfg_inputs["rotation_step"] = st.slider(
                        "Rotation Step (degrees)",
                        1,
                        90,
                        int(cfg_inputs.get("rotation_step", DEFAULT_ROT_STEP)),
                        key="rotation_step_slider",
                    )

                st.markdown("#### Efficiency")
                cfg_inputs["advanced_sort"] = st.toggle(
                    "Advanced Ordering",
                    value=bool(cfg_inputs.get("advanced_sort", True)),
                    help="Sort parts by area and aspect ratio before placement.",
                    key="adv_sort_toggle",
                )
                cfg_inputs["aggressive_packing"] = st.toggle(
                    "Aggressive Packing",
                    value=bool(cfg_inputs.get("aggressive_packing", True)),
                    help="Generate more candidate anchor positions (left/right/top combos).",
                    key="aggressive_packing_toggle",
                )
                cfg_inputs["enable_compaction"] = st.toggle(
                    "Post Compaction",
                    value=bool(cfg_inputs.get("enable_compaction", True)),
                    help="Slide parts left/up after initial placement to tighten gaps.",
                    key="compaction_toggle",
                )

            st.session_state.config_inputs = cfg_inputs

            st.write("")
            start_clicked = st.button(
                "Start Nesting",
                key="start_nesting_main",
                use_container_width=True,
            )
            if start_clicked:
                cfg_inputs = st.session_state.config_inputs
                sheet_w = float(cfg_inputs["sheet_w"])
                sheet_h = float(cfg_inputs["sheet_h"])
                spacing = float(cfg_inputs["spacing"])
                kerf = float(cfg_inputs["kerf"])
                rotation_mode = cfg_inputs["rotation_mode"]
                rotation_step = int(cfg_inputs.get("rotation_step", DEFAULT_ROT_STEP))

                all_parts = []
                for m in st.session_state.uploaded_meta:
                    if not m.get("valid"):
                        continue
                    qty = int(m.get("qty", 1))
                    for p in m.get("parts", []):
                        p_copy = dict(p)
                        p_copy["source_file"] = m["name"]
                        p_copy["label"] = m.get("label", m["name"])
                        for _ in range(qty):
                            all_parts.append(p_copy.copy())
                if not all_parts:
                    st.error("No valid parts to nest. Upload valid DXF files first.")
                else:
                    try:
                        with st.spinner("Nesting in progress..."):
                            if rotation_mode == "Free Rotation (Optimized)":
                                rotations = list(
                                    range(0, 360, max(1, 360 // FREE_ROTATION_SAMPLES))
                                )
                                free_rotation = True
                            else:
                                rotations = list(
                                    range(0, 360, max(1, rotation_step or 1))
                                )
                                free_rotation = False
                            sheets = nest_parts_improved(
                                all_parts,
                                sheet_w,
                                sheet_h,
                                spacing,
                                kerf,
                                rotations,
                                free_rotation=free_rotation,
                                advanced_sort=bool(cfg_inputs.get("advanced_sort", True)),
                                enable_compaction=bool(cfg_inputs.get("enable_compaction", True)),
                                aggressive_packing=bool(cfg_inputs.get("aggressive_packing", True)),
                            )
                            # Utilization metrics
                            def _sheet_util(sh):
                                area_sum = sum(calculate_bbox_area(p.get("placed_bbox", (0,0,0,0))) for p in sh)
                                sheet_area = sheet_w * sheet_h
                                return (area_sum / sheet_area * 100.0) if sheet_area > 0 else 0.0
                            utils = [_sheet_util(sh) for sh in sheets]
                            st.session_state.utilization_metrics = {
                                "per_sheet": utils,
                                "avg": sum(utils)/len(utils) if utils else 0.0,
                                "min": min(utils) if utils else 0.0,
                                "max": max(utils) if utils else 0.0,
                                "sheets": len(sheets),
                            }
                            st.session_state.nested_sheets = sheets
                            st.session_state.config = {
                                "sheet_w": sheet_w,
                                "sheet_h": sheet_h,
                                "spacing": spacing,
                                "kerf": kerf,
                                "rotation_mode": rotation_mode,
                            }
                            _auto_save_project()  # automatic save after successful nesting
                            st.success("Nesting complete")
                            st.toast("✅ Nesting completed", icon="✅")

                            part_rows = []
                            agg_counts = {}
                            for sheet in sheets:
                                for part in sheet:
                                    label = (
                                        part.get("label")
                                        or part.get("source_file")
                                        or "Unnamed part"
                                    )
                                    agg_counts[label] = agg_counts.get(label, 0) + 1
                            for lbl, qty_v in sorted(
                                agg_counts.items(), key=lambda x: x[0].lower()
                            ):
                                part_rows.append(
                                    {"Part Name": lbl, "Quantity": qty_v}
                                )
                            st.session_state.part_summary = part_rows

                        st.session_state.stage = "CUT"
                        # No explicit st.rerun(); state change triggers rerun.
                    except Exception as e:
                        st.error(f"Nesting failed: {e}")
                        st.exception(e)

    else:  # CUT stage
        st.markdown(
            "<div class='panel'><strong>Workspace & Results</strong></div>",
            unsafe_allow_html=True,
        )
        if not st.session_state.nested_sheets:
            st.info(
                "No nesting result yet. Go to **NEST** stage and run "
                "**Start Nesting**."
            )
        else:
            sheets = st.session_state.nested_sheets
            cfg_inputs = st.session_state.config_inputs
            cfg = st.session_state.config or {
                "sheet_w": float(cfg_inputs.get("sheet_w", DEFAULT_SHEET_W)),
                "sheet_h": float(cfg_inputs.get("sheet_h", DEFAULT_SHEET_H)),
                "spacing": float(cfg_inputs.get("spacing", DEFAULT_SPACING)),
                "kerf": float(cfg_inputs.get("kerf", DEFAULT_KERF)),
            }

            preview_canvas_px = int(cfg_inputs.get("preview_canvas_px", 900))
            show_bbox = bool(cfg_inputs.get("show_bbox", True))
            show_outline = bool(cfg_inputs.get("show_outline", True))

            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.metric("Sheets", len(sheets))
            with m2:
                total_parts_nested = sum(len(s) for s in sheets)
                st.metric("Total Parts", total_parts_nested)
            with m3:
                avg_parts = total_parts_nested / len(sheets) if sheets else 0
                st.metric("Avg Parts/Sheet", f"{avg_parts:.1f}")
            with m4:
                sheet_area = cfg["sheet_w"] * cfg["sheet_h"]
                total_area = sheet_area * len(sheets)
                st.metric("Total Area (m²)", f"{total_area / 1_000_000:.2f}")

            if (
                st.session_state.sheet_nav_number < 1
                or st.session_state.sheet_nav_number > len(sheets)
            ):
                st.session_state.sheet_nav_number = max(
                    1, min(len(sheets), st.session_state.sheet_nav_number)
                )

            def _step_sheet(delta: int, max_n: int):
                cur = int(st.session_state.get("sheet_nav_number", 1))
                st.session_state.sheet_nav_number = int(
                    max(1, min(max_n, cur + delta))
                )

            st.markdown("<div class='nav-area'>", unsafe_allow_html=True)
            nav_cols = st.columns([2, 2, 2, 3])
            with nav_cols[0]:
                st.number_input(
                    "Sheet #",
                    min_value=1,
                    max_value=len(sheets),
                    step=1,
                    key="sheet_nav_number",
                )
            with nav_cols[1]:
                st.button(
                    "Previous",
                    key="prev_btn",
                    use_container_width=True,
                    on_click=_step_sheet,
                    args=(-1, len(sheets)),
                )
            with nav_cols[2]:
                st.button(
                    "Next",
                    key="next_btn",
                    use_container_width=True,
                    on_click=_step_sheet,
                    args=(+1, len(sheets)),
                )
            with nav_cols[3]:
                st.write("")
                st.write(
                    f"Showing: {int(st.session_state.sheet_nav_number)}/{len(sheets)}"
                )
            st.markdown("</div>", unsafe_allow_html=True)

            sheet_idx = int(st.session_state.sheet_nav_number)
            selected = sheets[sheet_idx - 1]
            preview = create_sheet_preview_image(
                selected,
                cfg["sheet_w"],
                cfg["sheet_h"],
                canvas_px=preview_canvas_px,
                draw_bbox=show_bbox,
                draw_outline=show_outline,
            )
            if preview:
                st.image(
                    preview,
                    caption=f"Sheet {sheet_idx} preview",
                    use_container_width=True,
                )
            else:
                st.info("Preview unavailable")

            st.markdown("#### Sheet details")
            total_part_area = sum(
                calculate_bbox_area(p["placed_bbox"]) for p in selected
            )
            efficiency = (
                (total_part_area / (cfg["sheet_w"] * cfg["sheet_h"])) * 100
                if cfg["sheet_w"] * cfg["sheet_h"] > 0
                else 0
            )
            st.write(f"- Parts: **{len(selected)}**")
            st.write(
                f"- Sheet size: **{int(cfg['sheet_w'])} × {int(cfg['sheet_h'])} mm**"
            )
            st.write(f"- Utilization: **{efficiency:.1f}%**")

            if st.session_state.part_summary:
                st.markdown("#### Parts Summary (All Sheets)")
                rows = st.session_state.part_summary
                table_html = [
                    "<table style='width:100%; border-collapse:collapse;'>",
                    "<thead><tr>"
                    "<th style='text-align:left;padding:6px 8px;border-bottom:1px solid #334'>Part Name</th>"
                    "<th style='text-align:right;padding:6px 8px;border-bottom:1px solid #334'>Quantity</th>"
                    "</tr></thead><tbody>",
                ]
                for r in rows:
                    table_html.append(
                        "<tr>"
                        f"<td style='padding:4px 8px;border-bottom:1px solid #223'>{r['Part Name']}</td>"
                        f"<td style='padding:4px 8px;text-align:right;border-bottom:1px solid #223'>{r['Quantity']}</td>"
                        "</tr>"
                    )
                table_html.append("</tbody></table>")
                st.markdown("".join(table_html), unsafe_allow_html=True)

                col_copy, col_dl = st.columns([1, 1])
                with col_copy:
                    if st.button(
                        "Copy summary", key="copy_summary", use_container_width=True
                    ):
                        tsv = "Part Name\tQuantity\n" + "\n".join(
                            f"{r['Part Name']}\t{r['Quantity']}" for r in rows
                        )
                        st.code(tsv, language="")
                with col_dl:
                    if st.button(
                        "Download summary (.xls)",
                        key="dl_summary",
                        use_container_width=True,
                    ):
                        html = [
                            "<html><head><meta charset='utf-8'></head><body>"
                            "<table border='1'>",
                            "<tr><th>Part Name</th><th>Quantity</th></tr>",
                        ]
                        for r in rows:
                            html.append(
                                f"<tr><td>{r['Part Name']}</td>"
                                f"<td>{r['Quantity']}</td></tr>"
                            )
                        html.append("</table></body></html>")
                        data_bytes = "\n".join(html).encode("utf-8")
                        st.download_button(
                            "Download parts_summary.xls",
                            data=data_bytes,
                            file_name="parts_summary.xls",
                            mime="application/vnd.ms-excel",
                            use_container_width=True,
                        )

            st.markdown("#### Parts on sheet")
            cols_parts = st.columns(4)
            for i, p in enumerate(selected):
                thumb = Image.new("RGB", (240, 160), (255, 255, 255))
                d = ImageDraw.Draw(thumb)
                ob = p.get("orig_bbox") or (0, 0, 0, 0)
                if ob[2] - ob[0] > 0 and ob[3] - ob[1] > 0:
                    sx = 200 / (ob[2] - ob[0])
                    sy = 120 / (ob[3] - ob[1])
                    drew_any = False
                    outer = p.get("outer") or {}
                    pts_src = outer.get("points") if isinstance(outer, dict) else None
                    if pts_src:
                        pts = []
                        for pt in pts_src:
                            x = 10 + int((pt[0] - ob[0]) * sx)
                            y = 10 + int((ob[3] - pt[1]) * sy)
                            pts.append((x, y))
                        if len(pts) >= 2:
                            d.line(pts + [pts[0]], fill=(10, 90, 180), width=2)
                            drew_any = True
                    if not drew_any:
                        for ent in p.get("entities", []):
                            t = ent.get("type")
                            if t == "LWPOLYLINE" and ent.get("points"):
                                pts = []
                                for pt in ent["points"]:
                                    x = 10 + int((pt[0] - ob[0]) * sx)
                                    y = 10 + int((ob[3] - pt[1]) * sy)
                                    pts.append((x, y))
                                if len(pts) >= 2:
                                    d.line(
                                        pts + [pts[0]],
                                        fill=(50, 120, 200),
                                        width=2,
                                    )
                                    drew_any = True
                            elif t == "LINE":
                                s = ent.get("start")
                                e = ent.get("end")
                                if s and e:
                                    x1 = 10 + int((s[0] - ob[0]) * sx)
                                    y1 = 10 + int((ob[3] - s[1]) * sy)
                                    x2 = 10 + int((e[0] - ob[0]) * sx)
                                    y2 = 10 + int((ob[3] - e[1]) * sy)
                                    d.line(
                                        [(x1, y1), (x2, y2)],
                                        fill=(60, 130, 210),
                                        width=2,
                                    )
                                    drew_any = True
                        if not drew_any:
                            x1 = 10
                            y1 = 10
                            x2 = 10 + int((ob[2] - ob[0]) * sx)
                            y2 = 10 + int((ob[3] - ob[1]) * sy)
                            d.rectangle(
                                [x1, y1, x2, y2],
                                outline=(180, 190, 200),
                                width=2,
                            )
                with cols_parts[i % 4]:
                    rot = int(p.get("rotation", 0) or 0)
                    st.image(
                        thumb,
                        caption=f"Part {i+1} • {p.get('source_file','')} • {rot}°",
                        use_container_width=True,
                    )

            # Export section
            st.markdown("---")
            st.subheader("Export")

            fmt_col, scope_col = st.columns([2, 2])
            with fmt_col:
                export_format = st.radio(
                    "Format",
                    ["ZIP (.zip)", "DXF (.dxf)", "Excel (.xls)", "PDF (.pdf)"],
                    horizontal=True,
                    key="export_format",
                )
            with scope_col:
                if export_format in ("ZIP (.zip)", "Excel (.xls)", "PDF (.pdf)"):
                    export_scope = st.radio(
                        "Scope",
                        ["All sheets", "Single sheet"],
                        horizontal=True,
                        key="export_scope",
                    )
                else:
                    export_scope = "Single sheet"

            if export_scope == "Single sheet":
                st.number_input(
                    "Sheet #",
                    min_value=1,
                    max_value=len(sheets),
                    step=1,
                    key="export_sheet_number",
                )
                if (
                    st.session_state.export_sheet_number < 1
                    or st.session_state.export_sheet_number > len(sheets)
                ):
                    st.session_state.export_sheet_number = max(
                        1, min(len(sheets), st.session_state.export_sheet_number)
                    )
                sel_idx = int(st.session_state.export_sheet_number)
            else:
                sel_idx = 1

            if export_format == "ZIP (.zip)":
                btn_label = (
                    "Prepare ZIP (all sheets)"
                    if export_scope == "All sheets"
                    else f"Prepare ZIP (sheet {sel_idx})"
                )
                if st.button(
                    btn_label, key="zip_btn", use_container_width=True
                ):
                    try:
                        zip_buf = io.BytesIO()
                        with zipfile.ZipFile(
                            zip_buf, "w", zipfile.ZIP_DEFLATED
                        ) as zf:
                            temp_dir = tempfile.mkdtemp(prefix="ei8_nest_")
                            try:
                                base_out = os.path.join(temp_dir, "nested_out")
                                to_write = (
                                    sheets
                                    if export_scope == "All sheets"
                                    else [sheets[sel_idx - 1]]
                                )
                                write_sheets_entities_to_dxf(
                                    to_write,
                                    base_out,
                                    cfg["sheet_w"],
                                    cfg["sheet_h"],
                                )
                                for fname in sorted(os.listdir(temp_dir)):
                                    if fname.lower().endswith(".dxf"):
                                        path = os.path.join(temp_dir, fname)
                                        with open(path, "rb") as fh:
                                            zf.writestr(fname, fh.read())
                            finally:
                                try:
                                    for fname in os.listdir(temp_dir):
                                        os.remove(os.path.join(temp_dir, fname))
                                    os.rmdir(temp_dir)
                                except Exception:
                                    pass
                        zip_buf.seek(0)
                        zip_name = (
                            "nested_sheets.zip"
                            if export_scope == "All sheets"
                            else f"nested_sheet{sel_idx:03d}.zip"
                        )
                        st.download_button(
                            f"Download {zip_name}",
                            data=zip_buf.getvalue(),
                            file_name=zip_name,
                            mime="application/zip",
                            use_container_width=True,
                        )
                        st.toast("📦 ZIP ready", icon="📦")
                    except Exception as e:
                        st.error(f"Export failed: {e}")
                        st.exception(e)

            elif export_format == "DXF (.dxf)":
                if st.button(
                    f"Prepare DXF (sheet {sel_idx})",
                    key="dxf_btn",
                    use_container_width=True,
                ):
                    try:
                        temp_dir = tempfile.mkdtemp(prefix="ei8_nest_")
                        try:
                            base_out = os.path.join(temp_dir, "single_out")
                            write_sheets_entities_to_dxf(
                                [sheets[sel_idx - 1]],
                                base_out,
                                cfg["sheet_w"],
                                cfg["sheet_h"],
                            )
                            dxf_path = None
                            for fname in os.listdir(temp_dir):
                                if fname.lower().endswith(".dxf"):
                                    dxf_path = os.path.join(temp_dir, fname)
                                    break
                            if not dxf_path:
                                raise RuntimeError("DXF not generated")
                            with open(dxf_path, "rb") as fh:
                                data_bytes = fh.read()
                            out_name = f"nested_sheet{sel_idx:03d}.dxf"
                            st.download_button(
                                f"Download {out_name}",
                                data=data_bytes,
                                file_name=out_name,
                                mime="application/octet-stream",
                                use_container_width=True,
                            )
                            st.toast("📄 DXF ready", icon="📄")
                        finally:
                            try:
                                for fname in os.listdir(temp_dir):
                                    os.remove(os.path.join(temp_dir, fname))
                                os.rmdir(temp_dir)
                            except Exception:
                                pass
                    except Exception as e:
                        st.error(f"Export failed: {e}")
                        st.exception(e)
            elif export_format == "PDF (.pdf)":
                # No interactive options: always generate detailed report layout
                use_report = True
                title_text = "Optimization Nesting Report"
                logo_file = None

                btn_label = (
                    "Prepare PDF (all sheets)"
                    if export_scope == "All sheets"
                    else f"Prepare PDF (sheet {sel_idx})"
                )
                if st.button(btn_label, key="pdf_btn", use_container_width=True):
                    try:
                        pdf_buf = io.BytesIO()
                        did_report = False
                        if use_report and REPORTLAB_AVAILABLE:
                            # Try to load a default logo from disk; fallback to text logo
                            logo_bytes = None
                            for cand in ("logo.png", "assets/logo.png", "ei8.png", "assets/ei8.png"):
                                p = os.path.join(os.getcwd(), cand)
                                if os.path.exists(p):
                                    try:
                                        with open(p, "rb") as fh:
                                            logo_bytes = fh.read()
                                        break
                                    except Exception:
                                        pass
                            data = build_pdf_report(
                                sheets,
                                cfg,
                                preview_canvas_px,
                                show_bbox,
                                show_outline,
                                scope=export_scope,
                                sel_idx=sel_idx,
                                uploaded_meta=st.session_state.uploaded_meta,
                                config_inputs=st.session_state.config_inputs,
                                logo_bytes=logo_bytes,
                                title_text=title_text,
                            )
                            st.download_button(
                                "Download report.pdf" if export_scope == "All sheets" else f"Download report_sheet{sel_idx:03d}.pdf",
                                data=data,
                                file_name=("report.pdf" if export_scope == "All sheets" else f"report_sheet{sel_idx:03d}.pdf"),
                                mime="application/pdf",
                                use_container_width=True,
                            )
                            st.toast("📄 PDF ready", icon="📄")
                            did_report = True
                        elif use_report and not REPORTLAB_AVAILABLE:
                            st.warning("ReportLab not installed. Falling back to image-only PDF.")

                        # Fallback: existing image-based PDF export
                        if not did_report and export_scope == "All sheets":
                            images = []
                            for sh in sheets:
                                img = create_sheet_preview_image(
                                    sh,
                                    cfg["sheet_w"],
                                    cfg["sheet_h"],
                                    canvas_px=preview_canvas_px,
                                    draw_bbox=show_bbox,
                                    draw_outline=show_outline,
                                )
                                if img is None:
                                    continue
                                if img.mode != "RGB":
                                    img = img.convert("RGB")
                                images.append(img)
                            if not images:
                                st.error("No previews available to export.")
                            else:
                                first, rest = images[0], images[1:]
                                first.save(
                                    pdf_buf,
                                    format="PDF",
                                    save_all=True,
                                    append_images=rest,
                                )
                                pdf_buf.seek(0)
                                st.download_button(
                                    "Download nesting_all_sheets.pdf",
                                    data=pdf_buf.getvalue(),
                                    file_name="nesting_all_sheets.pdf",
                                    mime="application/pdf",
                                    use_container_width=True,
                                )
                                st.toast("📄 PDF ready", icon="📄")
                        elif not did_report:  # single sheet
                            sh = sheets[sel_idx - 1]
                            img = create_sheet_preview_image(
                                sh,
                                cfg["sheet_w"],
                                cfg["sheet_h"],
                                canvas_px=preview_canvas_px,
                                draw_bbox=show_bbox,
                                draw_outline=show_outline,
                            )
                            if img is None:
                                st.error("Preview generation failed for PDF export.")
                            else:
                                if img.mode != "RGB":
                                    img = img.convert("RGB")
                                img.save(pdf_buf, format="PDF")
                                pdf_buf.seek(0)
                                out_name = f"nested_sheet{sel_idx:03d}.pdf"
                                st.download_button(
                                    f"Download {out_name}",
                                    data=pdf_buf.getvalue(),
                                    file_name=out_name,
                                    mime="application/pdf",
                                    use_container_width=True,
                                )
                                st.toast("📄 PDF ready", icon="📄")
                    except Exception as e:
                        st.error(f"Export failed: {e}")
                        st.exception(e)
            else:  # Excel
                btn_label = (
                    "Prepare XLS (all sheets)"
                    if export_scope == "All sheets"
                    else f"Prepare XLS (sheet {sel_idx})"
                )
                if st.button(
                    btn_label, key="xls_btn", use_container_width=True
                ):
                    try:
                        target_sheets = (
                            sheets
                            if export_scope == "All sheets"
                            else [sheets[sel_idx - 1]]
                        )
                        rows = []
                        for s_idx, sheet in enumerate(
                            target_sheets,
                            start=(1 if export_scope == "All sheets" else sel_idx),
                        ):
                            sheet_area = cfg["sheet_w"] * cfg["sheet_h"]
                            part_area_sum = sum(
                                calculate_bbox_area(
                                    p.get("placed_bbox", (0, 0, 0, 0))
                                )
                                for p in sheet
                            )
                            utilization = (
                                (part_area_sum / sheet_area) * 100
                                if sheet_area > 0
                                else 0
                            )
                            for p_idx, part in enumerate(sheet, start=1):
                                bbox = part.get("placed_bbox", (0, 0, 0, 0))
                                rows.append(
                                    [
                                        s_idx,
                                        p_idx,
                                        part.get("source_file", ""),
                                        part.get("label", ""),
                                        part.get("rotation", 0),
                                        part.get("anchor_X", 0),
                                        part.get("anchor_Y", 0),
                                        bbox[0],
                                        bbox[1],
                                        bbox[2],
                                        bbox[3],
                                        (bbox[2] - bbox[0]),
                                        (bbox[3] - bbox[1]),
                                        round(utilization, 2),
                                    ]
                                )
                        headers = [
                            "sheet",
                            "part_index",
                            "source_file",
                            "label",
                            "rotation_deg",
                            "anchor_x",
                            "anchor_y",
                            "bbox_x1",
                            "bbox_y1",
                            "bbox_x2",
                            "bbox_y2",
                            "bbox_w",
                            "bbox_h",
                            "sheet_utilization_%",
                        ]

                        def _esc(v):
                            s = str(v)
                            return (
                                s.replace("&", "&amp;")
                                .replace("<", "&lt;")
                                .replace(">", "&gt;")
                            )

                        html = [
                            "<html><head><meta charset='utf-8'></head><body><table border='1'>"
                        ]
                        html.append(
                            "<tr>"
                            + "".join(f"<th>{_esc(h)}</th>" for h in headers)
                            + "</tr>"
                        )
                        for r in rows:
                            html.append(
                                "<tr>"
                                + "".join(f"<td>{_esc(c)}</td>" for c in r)
                                + "</tr>"
                            )
                        html.append("</table></body></html>")
                        data_bytes = "\n".join(html).encode("utf-8")
                        fname = (
                            "nested_summary.xls"
                            if export_scope == "All sheets"
                            else f"nested_sheet{sel_idx:03d}_summary.xls"
                        )
                        st.download_button(
                            f"Download {fname}",
                            data=data_bytes,
                            file_name=fname,
                            mime="application/vnd.ms-excel",
                            use_container_width=True,
                        )
                        st.toast("📊 XLS ready", icon="📊")
                    except Exception as e:
                        st.error(f"Export failed: {e}")
                        st.exception(e)

with col_r:
    st.markdown(
        "<div class='panel'><strong>Stage & Tips</strong></div>",
        unsafe_allow_html=True,
    )
    if stage == "SET":
        st.info(
            "SET: Upload DXF files, adjust labels and quantities. "
            "Sheet & nesting parameters are configured in the NEST stage."
        )
    elif stage == "NEST":
        st.info(
            "NEST: Configure sheet and nesting parameters, then run "
            "**Start Nesting** to generate layouts."
        )
    else:
        st.info(
            "CUT: Review nesting results, sheet utilization and part summaries, "
            "then export DXF / ZIP / XLS as needed."
        )

prev_map = {"SET": None, "NEST": "SET", "CUT": "NEST"}
next_map = {"SET": "NEST", "NEST": "CUT", "CUT": None}

c_prev, c_next, _ = st.columns([1, 1, 6])
with c_prev:
    prev_stage = prev_map[stage]
    if prev_stage:
        st.button(
            "Back",
            key=f"back_{stage}",
            use_container_width=True,
            on_click=go_to_stage,
            args=(prev_stage,),
        )
    else:
        st.button(
            "Back",
            key=f"back_{stage}",
            use_container_width=True,
            disabled=True,
        )

with c_next:
    next_stage = next_map[stage]
    disabled_next = False
    if not next_stage:
        disabled_next = True
    elif stage == "NEST" and not st.session_state.get("nested_sheets"):
        disabled_next = True

    if next_stage:
        st.button(
            "Next",
            key=f"next_{stage}",
            use_container_width=True,
            disabled=disabled_next,
            on_click=go_to_stage,
            args=(next_stage,),
        )
    else:
        st.button(
            "Next",
            key=f"next_{stage}",
            use_container_width=True,
            disabled=True,
        )


