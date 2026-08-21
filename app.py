"""
app.py

Flask web front-end for the AMBR plate-mapping tool, meant to be deployed to
a host like Render.com where there's no desktop/display available (so the
tkinter GUI from generate_plate_mapping.py can't run). The same processing
logic lives in core.py, imported here unchanged; this file only replaces the
tkinter dialogs with a multi-step web form:

  1. Upload AMBR timepoints.csv + Benchling Timepoint-Sample.csv, and
     optionally a sampling_scheme.xlsx.
  2. Confirm which sample volumes to exclude (shown with row counts, exactly
     like the desktop checkbox popup).
  3. Confirm the source plate format (24- or 96-well) -- same reasoning as
     the desktop version: well usage can only prove a lower bound, so the
     user is always asked, with genuine contradictions (e.g. a well like F8)
     overriding whatever was picked.
  4. Choose whether/how to repack samples onto an analysis plate for
     HPLC/SOA (pooling every source plate into one continuous run first).
  5. Download the resulting file(s).

State between steps is kept server-side, in a per-visitor temp directory
keyed by a random id stored in the Flask session cookie -- nothing is
shared between different users hitting the same deployment.
"""

import os
import pickle
import secrets
import tempfile
import time

from flask import Flask, flash, redirect, render_template, request, send_file, session, url_for
from werkzeug.utils import secure_filename

import core

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB of uploads per request, plenty for these files

BASE_TMP = os.path.join(tempfile.gettempdir(), "ambr_plate_mapping_sessions")
os.makedirs(BASE_TMP, exist_ok=True)
SESSION_MAX_AGE_SECONDS = 6 * 3600  # stale session dirs older than this are cleaned up opportunistically


def _session_dir():
    sid = session.get("sid")
    if not sid:
        sid = secrets.token_urlsafe(16)
        session["sid"] = sid
    d = os.path.join(BASE_TMP, sid)
    os.makedirs(d, exist_ok=True)
    return d


def _state_path():
    return os.path.join(_session_dir(), "state.pkl")


def _load_state():
    path = _state_path()
    if not os.path.exists(path):
        return {}
    with open(path, "rb") as f:
        return pickle.load(f)


def _save_state(state):
    with open(_state_path(), "wb") as f:
        pickle.dump(state, f)


def _cleanup_stale_sessions():
    """Best-effort cleanup of old per-visitor temp directories so a
    long-running Render instance doesn't slowly fill its disk. Not
    security-critical -- Render's free/standard disks are ephemeral anyway
    and get wiped on redeploy, this just keeps things tidy in between."""
    try:
        now = time.time()
        for name in os.listdir(BASE_TMP):
            p = os.path.join(BASE_TMP, name)
            try:
                if os.path.isdir(p) and (now - os.path.getmtime(p)) > SESSION_MAX_AGE_SECONDS:
                    for root, dirs, files in os.walk(p, topdown=False):
                        for fn in files:
                            os.remove(os.path.join(root, fn))
                        for dn in dirs:
                            os.rmdir(os.path.join(root, dn))
                    os.rmdir(p)
            except OSError:
                continue
    except OSError:
        pass


@app.route("/", methods=["GET"])
def index():
    _cleanup_stale_sessions()
    session.pop("sid", None)  # starting over always gets a fresh session dir
    return render_template("upload.html")


@app.route("/upload", methods=["POST"])
def upload():
    d = _session_dir()

    timepoints_file = request.files.get("timepoints")
    benchling_file = request.files.get("benchling_template")
    scheme_file = request.files.get("scheme")

    if not timepoints_file or not timepoints_file.filename:
        flash("Please choose an AMBR timepoints .csv file.")
        return redirect(url_for("index"))
    if not benchling_file or not benchling_file.filename:
        flash("Please choose a Benchling Timepoint-Sample .csv file.")
        return redirect(url_for("index"))

    timepoints_path = os.path.join(d, "timepoints_" + secure_filename(timepoints_file.filename))
    timepoints_file.save(timepoints_path)

    benchling_path = os.path.join(d, "benchling_" + secure_filename(benchling_file.filename))
    benchling_file.save(benchling_path)

    scheme_path = None
    if scheme_file and scheme_file.filename:
        scheme_path = os.path.join(d, "scheme_" + secure_filename(scheme_file.filename))
        scheme_file.save(scheme_path)

    state = {
        "timepoints_path": timepoints_path,
        "benchling_path": benchling_path,
        "scheme_path": scheme_path,
    }
    _save_state(state)
    return redirect(url_for("volumes"))


@app.route("/volumes", methods=["GET"])
def volumes():
    state = _load_state()
    if "timepoints_path" not in state:
        return redirect(url_for("index"))
    counts = core.scan_volumes(state["timepoints_path"])
    sorted_vols = sorted(counts, key=lambda v: float(v)) if counts else []
    return render_template("volumes.html", volumes=sorted_vols, counts=counts)


@app.route("/volumes", methods=["POST"])
def volumes_submit():
    state = _load_state()
    if "timepoints_path" not in state:
        return redirect(url_for("index"))

    exclude_volumes = request.form.getlist("exclude_volume")
    exclude_text_raw = request.form.get("exclude_text", "END SAMPLES").strip()
    exclude_text = [t.strip() for t in exclude_text_raw.split(",") if t.strip()] or ["END SAMPLES"]

    events, dropped = core.parse_timepoints(state["timepoints_path"], exclude_text, exclude_volumes)
    experiment_name, reactor_info = core.parse_benchling_template(state["benchling_path"])

    if state.get("scheme_path"):
        scheme = core.parse_scheme(state["scheme_path"])
        scheme_source = "file"
        generated_scheme_path = None
    else:
        scheme = core.derive_scheme_from_events(events)
        scheme_source = "auto"
        generated_scheme_path = os.path.join(_session_dir(), f"{experiment_name or 'AMBR'}_sampling_scheme.xlsx")
        core.write_scheme_workbook(events, generated_scheme_path)

    rows, unused_slots = core.build_rows(events, scheme, experiment_name, reactor_info, scheme_source=scheme_source)
    detection = core.detect_plate_format(rows)

    state.update(
        {
            "exclude_volumes": exclude_volumes,
            "exclude_text": exclude_text,
            "experiment_name": experiment_name,
            "scheme_source": scheme_source,
            "generated_scheme_path": generated_scheme_path,
            "rows": rows,
            "unused_slots": unused_slots,
            "dropped": dropped,
            "detection": detection,
        }
    )
    _save_state(state)
    return redirect(url_for("plate_format"))


@app.route("/plate_format", methods=["GET"])
def plate_format():
    state = _load_state()
    if "rows" not in state:
        return redirect(url_for("index"))
    return render_template("plate_format.html", detection=state["detection"])


@app.route("/plate_format", methods=["POST"])
def plate_format_submit():
    state = _load_state()
    if "rows" not in state:
        return redirect(url_for("index"))

    src_fmt = request.form.get("source_plate_format", "24")
    detection = state["detection"]
    override_note = None
    if src_fmt == "24" and detection["format"] == "96":
        override_note = (
            f"You said 24-well, but well '{detection['max_row_letter']}{detection['max_col']}' appears in the "
            "data, which is impossible on a 24-well plate (max is D6). Using 96-well instead, since the data proves it."
        )
        src_fmt = "96"

    state["src_fmt"] = src_fmt
    state["override_note"] = override_note
    _save_state(state)
    return redirect(url_for("transpose_step"))


@app.route("/transpose", methods=["GET"])
def transpose_step():
    state = _load_state()
    if "src_fmt" not in state:
        return redirect(url_for("index"))
    other_fmt = "96" if state["src_fmt"] == "24" else "24"
    return render_template(
        "transpose.html", src_fmt=state["src_fmt"], other_fmt=other_fmt, override_note=state.get("override_note")
    )


@app.route("/transpose", methods=["POST"])
def transpose_submit():
    state = _load_state()
    if "src_fmt" not in state:
        return redirect(url_for("index"))

    transpose_to = request.form.get("transpose_to", "none")
    d = _session_dir()

    output_path = os.path.join(d, "plate_mapping_output.xlsx")
    core.write_workbook(state["rows"], state["unused_slots"], state["experiment_name"], output_path)

    transpose_output_path = None
    n_dest_plates = None
    if transpose_to in ("24", "96"):
        transposed_rows = core.build_transposition(state["rows"], transpose_to)
        transpose_output_path = os.path.join(
            d, f"plate_mapping_output_transposed_to_{transpose_to}well.xlsx"
        )
        core.write_transposed_workbook(transposed_rows, state["src_fmt"], transpose_to, transpose_output_path)
        n_dest_plates = len({r["dest_plate_label"] for r in transposed_rows})

    not_in_scheme = sum(1 for r in state["rows"] if r["in_scheme"].startswith("NO"))

    state.update(
        {
            "output_path": output_path,
            "transpose_output_path": transpose_output_path,
            "transpose_to": transpose_to,
            "n_dest_plates": n_dest_plates,
            "not_in_scheme": not_in_scheme,
        }
    )
    _save_state(state)
    return redirect(url_for("done"))


@app.route("/done", methods=["GET"])
def done():
    state = _load_state()
    if "output_path" not in state:
        return redirect(url_for("index"))
    return render_template(
        "done.html",
        experiment_name=state.get("experiment_name"),
        total_rows=len(state["rows"]),
        dropped_count=len(state["dropped"]),
        not_in_scheme=state["not_in_scheme"],
        unused_count=len(state["unused_slots"]),
        scheme_source=state["scheme_source"],
        has_generated_scheme=bool(state.get("generated_scheme_path")),
        transpose_to=state.get("transpose_to"),
        n_dest_plates=state.get("n_dest_plates"),
    )


@app.route("/download/output")
def download_output():
    state = _load_state()
    return send_file(state["output_path"], as_attachment=True, download_name="plate_mapping_output.xlsx")


@app.route("/download/scheme")
def download_scheme():
    state = _load_state()
    path = state.get("generated_scheme_path")
    if not path:
        return redirect(url_for("done"))
    name = os.path.basename(path)
    return send_file(path, as_attachment=True, download_name=name)


@app.route("/download/transposed")
def download_transposed():
    state = _load_state()
    path = state.get("transpose_output_path")
    if not path:
        return redirect(url_for("done"))
    fmt = state.get("transpose_to", "")
    return send_file(path, as_attachment=True, download_name=f"plate_mapping_output_transposed_to_{fmt}well.xlsx")


if __name__ == "__main__":
    app.run(debug=True, port=int(os.environ.get("PORT", 5000)))
