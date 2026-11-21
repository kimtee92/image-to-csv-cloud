#!/usr/bin/env bash
#
# Deploy image-to-csv to Cloud Run.
#
# Builds from a temporary staging directory containing only the application files.
# `gcloud run deploy --source .` uploads the whole working tree to a Cloud Storage
# staging bucket, and .dockerignore does NOT prevent that — it only controls what
# lands in the image. Staging keeps client scans in reference/ and reference2/ out
# of the upload.
#
# The staging tree mirrors the repo layout (app/...), so the real Dockerfile is used
# exactly as-is instead of a second inline copy that could silently drift out of sync.
#
# Tunables, all optional — export them before running:
#   SERVICE REGION MEMORY CPU CONCURRENCY MIN_INSTANCES MAX_INSTANCES TIMEOUT
#   GEMINI_MODEL REASONING_EFFORT DTR_CROSSCHECK IMAGE_PREPROCESS TIME_VERIFY
#   OCR_TIMEOUT VERIFY_TIMEOUT
#   GEMINI_SECRET   name of a Secret Manager secret holding the API key
#   DRY_RUN=1       print the gcloud command instead of running it
set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SERVICE="${SERVICE:-image-to-csv}"
REGION="${REGION:-us-central1}"

# 1Gi, not the old default: a DTR re-read decodes the source scan and builds a
# magnified band crop (capped at 8M pixels) on top of the upscaled model copy, so
# peak RSS per request is far higher than it was for a single-pass extraction.
# Concurrency is set with that in mind — raise memory before raising concurrency.
MEMORY="${MEMORY:-1Gi}"
CPU="${CPU:-1}"
CONCURRENCY="${CONCURRENCY:-4}"
MIN_INSTANCES="${MIN_INSTANCES:-0}"
MAX_INSTANCES="${MAX_INSTANCES:-3}"
# OCR_TIMEOUT (170) + VERIFY_TIMEOUT (90) = 260s of model calls must finish inside
# the Cloud Run request timeout, or a slow pair of calls is dropped with no result.
TIMEOUT="${TIMEOUT:-300}"

STAGE="$(mktemp -d)"
# Runs on success, on failure, and on Ctrl-C. With `set -e` a plain cleanup line at
# the end of the script would be skipped whenever anything above it failed.
trap 'rm -rf "$STAGE"' EXIT

mkdir -p "$STAGE/app"
cp "$REPO/Dockerfile" "$STAGE/"
cp "$REPO/app/requirements.txt" "$STAGE/app/"
cp -r "$REPO/app/static" "$STAGE/app/"
cp -r "$REPO/app/templates" "$STAGE/app/"
# Copy every module rather than naming them one by one. main.py imports its
# siblings as top-level modules (the container has them all in /app), so a module
# missing from the stage still builds a perfectly good image that then dies on
# startup with ImportError. Globbing removes that failure mode for good.
cp "$REPO"/app/*.py "$STAGE/app/"

echo "Staged for deploy:"
find "$STAGE" -type f -printf '  %P\n' | sort

# ── preflight ────────────────────────────────────────────────────────────────
# Catch here what would otherwise cost a full Cloud Build round trip, or worse,
# deploy green and fail on the first request.
fail() { echo "deploy aborted: $*" >&2; exit 1; }

[[ -f "$STAGE/app/main.py" ]] || fail "app/main.py is missing from the stage"
[[ -f "$STAGE/app/static/index.html" ]] || fail "app/static/index.html is missing"
[[ -d "$STAGE/app/templates" ]] || fail "app/templates is missing"
tpl_n=$(ls "$STAGE/app/templates"/*.xlsx 2>/dev/null | wc -l)
(( tpl_n > 0 )) || fail "no client workbook templates staged"
echo "  client templates staged: $tpl_n"

# Compile against the Python the CONTAINER runs, not the one on this machine.
# These drift: a backslash inside an f-string expression is legal from 3.12 but a
# SyntaxError on 3.11, so a local compileall passes and Cloud Run then fails with
# "container failed to start and listen on PORT". Docker gives us the real answer;
# without it, say so rather than implying the check was thorough.
BASE_IMAGE="$(awk '/^FROM /{print $2; exit}' "$REPO/Dockerfile")"
if command -v docker >/dev/null && docker info >/dev/null 2>&1; then
  echo "Checking syntax under $BASE_IMAGE ..."
  docker run --rm -v "$STAGE/app:/src:ro" "$BASE_IMAGE" python -c '
import pathlib, sys
bad = 0
for f in sorted(pathlib.Path("/src").glob("*.py")):
    try:
        compile(f.read_text(), f.name, "exec")
    except SyntaxError as e:
        bad = 1
        print(f"  {f.name}: SyntaxError line {e.lineno}: {e.msg}", file=sys.stderr)
sys.exit(bad)
' || fail "staged code does not compile on $BASE_IMAGE"
else
  local_py="$(python3 -c 'import sys; print("%d.%d" % sys.version_info[:2])' 2>/dev/null || echo "?")"
  base_py="$(sed -E 's/^python:([0-9]+\.[0-9]+).*/\1/' <<<"$BASE_IMAGE")"
  if [[ "$local_py" != "$base_py" ]]; then
    echo "  WARNING: docker unavailable; syntax checked with local Python $local_py" >&2
    echo "           but the container runs $base_py — version-specific syntax will" >&2
    echo "           not be caught here." >&2
  fi
fi

if command -v python3 >/dev/null; then
  python3 -m compileall -q "$STAGE/app" >/dev/null \
    || fail "a staged Python file does not compile"

  # Every top-level import in main.py has to resolve to something the container
  # will actually have: the standard library, a pinned dependency, or a module
  # sitting beside it in /app. An earlier version checked the repo first and so
  # treated a deleted module as third-party — exactly backwards, since a module
  # that exists nowhere is the failure worth catching.
  python3 - "$STAGE/app" <<'PY' || fail "unresolved import (see above)"
import ast, pathlib, sys, sysconfig

stage = pathlib.Path(sys.argv[1])
# Scan every module, not just main.py: dtr_export imports xlsx_writer, so checking
# only the entrypoint misses a sibling that has gone missing one level down.
# Distribution name -> module name, for the few that differ.
ALIASES = {"pillow": "PIL", "google-auth": "google", "python-multipart": "multipart"}
deps = {"pydantic"}          # transitive via fastapi, imported directly
for line in (stage / "requirements.txt").read_text().splitlines():
    name = line.split("==")[0].split(">=")[0].split("[")[0].strip().lower()
    if name and not name.startswith("#"):
        deps.add(ALIASES.get(name, name).lower())

roots: dict[str, str] = {}          # module root -> the file that imports it
for path in sorted(stage.glob("*.py")):
    for node in ast.walk(ast.parse(path.read_text())):
        if isinstance(node, ast.Import):
            for alias in node.names:
                roots.setdefault(alias.name.split(".")[0], path.name)
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            roots.setdefault(node.module.split(".")[0], path.name)

stdlib = set(getattr(sys, "stdlib_module_names", ()))
bad = [
    (root, src) for root, src in sorted(roots.items())
    if root not in stdlib and root.lower() not in deps
    and not (stage / f"{root}.py").exists()
]
for root, src in bad:
    print(f"  {src} imports '{root}', which is neither stdlib, nor in "
          f"requirements.txt, nor a module in app/", file=sys.stderr)
sys.exit(1 if bad else 0)
PY

  # The glob above should make this impossible; assert it rather than assume it.
  diff <(cd "$REPO/app" && ls *.py | sort) <(cd "$STAGE/app" && ls *.py | sort) >/dev/null \
    || fail "staged modules differ from app/*.py"
fi
# Every client workbook must survive a fill and come out as valid OOXML. Excel
# refuses a file whose mc:Ignorable names a namespace prefix the part no longer
# declares, and ElementTree drops exactly those declarations when it re-serialises
# a sheet. LibreOffice opens such a file happily, so nothing here caught it until
# a client hit "We found a problem with some content".
#
# Each template is filled twice: a small batch, and one past the template's own
# block count so the cloning path runs. The first version of this check only did
# the small batch, and so missed cloned shared-formula dependents pointing at a
# group whose range did not cover them — Excel repaired those files and wiped
# every block from its 3rd row down to #N/A.
if command -v python3 >/dev/null; then
  echo "Checking generated workbooks open as valid OOXML ..."
  PYTHONPATH="$STAGE/app" python3 - "$STAGE/app" <<'PY' || fail "generated workbook is not valid OOXML (see above)"
import re, sys, zipfile, io, asyncio, pathlib
import xml.etree.ElementTree as ET
sys.path.insert(0, sys.argv[1])
from clients import CLIENTS, LAYOUTS, TEMPLATES
import formula_compat
import html
import xlsx_fill

bad = []
for key, client in CLIENTS.items():
    tpl = client.get("template")
    if not tpl:
        continue
    cfg = TEMPLATES[tpl]
    slots = LAYOUTS[client["layout"]]["slots"]
    rows = [{s: ("07:30" if s.endswith("_in") else "17:00" if s.endswith("_out") else "")
             for s in slots} for _ in range(2)]
    path = str(pathlib.Path(sys.argv[1]) / "templates" / cfg["file"])
    for size in (3, int(cfg.get("blocks") or 1) + 2):
        blocks = [{"name": "PREFLIGHT %d" % i, "rows": rows} for i in range(1, size + 1)]
        label = f"{key} x{size}"
        try:
            content, _ = xlsx_fill.fill(path, blocks, cfg)
        except Exception as exc:
            bad.append(f"  {label}: fill failed: {exc}")
            continue
        z = zipfile.ZipFile(io.BytesIO(content))
        if z.testzip() is not None:
            bad.append(f"  {label}: corrupt zip entry")
        types = z.read("[Content_Types].xml").decode("utf-8", "ignore")
        for name in z.namelist():
            if not name.endswith((".xml", ".rels")):
                continue
            raw = z.read(name)
            try:
                root = ET.fromstring(raw)
            except Exception as exc:
                bad.append(f"  {label}/{name}: malformed XML ({exc})")
                continue
            text = raw.decode("utf-8", "ignore")
            # The export forces a full recalculation, so a function Excel 2016/2019
            # lacks shows #NAME? there (24 Chicken ND column, Nov 2026). Every
            # LET/SWITCH/XLOOKUP must have been rewritten by formula_compat.
            if name.startswith("xl/worksheets/"):
                left = set()
                for fm in re.finditer(r"<f\b[^>]*>([^<]*)</f>", text):
                    left |= formula_compat.unsupported_functions(html.unescape(fm.group(1)))
                if left:
                    bad.append(f"  {label}/{name}: formulas still need Excel 365: {sorted(left)}")
            m = re.search(r'Ignorable="([^"]*)"', text)
            if m:
                declared = set(re.findall(r"xmlns:([A-Za-z0-9_]+)=", text))
                missing = [p for p in m.group(1).split() if p not in declared]
                if missing:
                    bad.append(f"  {label}/{name}: mc:Ignorable names undeclared {missing}")
            if name.startswith("xl/worksheets/sheet") and f"/{name}" not in types:
                bad.append(f"  {label}/{name}: no content-type override")
            if not name.startswith("xl/worksheets/sheet"):
                continue
            # Shared formulas: every dependent must sit inside its master's range.
            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            masters, deps = {}, []
            for c in root.iter(ns + "c"):
                f = c.find(ns + "f")
                if f is None or f.get("t") != "shared":
                    continue
                if f.get("ref"):
                    masters[f.get("si")] = f.get("ref")
                else:
                    deps.append((c.get("r"), f.get("si")))
            def pos(ref):
                mm = re.match(r"([A-Z]+)(\d+)$", ref)
                col = 0
                for ch in mm.group(1):
                    col = col * 26 + ord(ch) - 64
                return col, int(mm.group(2))
            orphans = 0
            for cell, si in deps:
                rng = masters.get(si)
                if rng is None:
                    orphans += 1
                    continue
                a, b = (rng.split(":") + [rng])[:2]
                (c1, r1), (c2, r2), (cc, rr) = pos(a), pos(b), pos(cell)
                if not (c1 <= cc <= c2 and r1 <= rr <= r2):
                    orphans += 1
            if orphans:
                bad.append(f"  {label}/{name}: {orphans} shared-formula cell(s) "
                           f"outside their group (Excel will repair to #N/A)")
        # Formula holes: a computed column whose formula is missing on some data
        # rows. The clients' own working files carry these (Bikezilla had 13 FINAL
        # OUT cells typed over), and a day landing on one computes 0 hours.
        sheet_xml = z.read(xlsx_fill._sheet_path(z, cfg["sheet"]))
        froot = ET.fromstring(sheet_xml)
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        has = {c.get("r") for c in froot.iter(ns + "c") if c.find(ns + "f") is not None}
        ref_row = cfg["first_block"] + 2
        fcols = {re.match(r"[A-Z]+", r).group() for r in has
                 if re.fullmatch(r"[A-Z]+" + str(ref_row), r)} - set(cfg["columns"])
        holes = []
        for b in range(max(size, int(cfg.get("blocks") or 1))):
            rows_b = [cfg["first_block"] + cfg["stride"] * b + 2 + k for k in range(cfg["rows_per_block"])]
            for col in fcols:
                if sum(f"{col}{r}" in has for r in rows_b) >= cfg["rows_per_block"] // 2:
                    holes += [f"{col}{r}" for r in rows_b if f"{col}{r}" not in has]
        if holes:
            bad.append(f"  {label}: {len(holes)} formula hole(s) in computed columns "
                       f"(days there compute 0 hours): {holes[:6]}")

if bad:
    print("workbook validation failed:", file=sys.stderr)
    for line in bad:
        print(line, file=sys.stderr)
    sys.exit(1)
print("  all client workbooks valid")
PY
fi

echo "Preflight OK"

# Cloud Build runs this script with PREFLIGHT_ONLY=1 as its first step, so pushes
# to GitHub get exactly the same checks as a manual deploy before anything ships.
if [[ -n "${PREFLIGHT_ONLY:-}" ]]; then
  exit 0
fi

# ── env vars ─────────────────────────────────────────────────────────────────
# Only forward the settings actually exported for this run, and use
# --update-env-vars rather than --set-env-vars: the latter REPLACES the service's
# whole environment, silently dropping anything configured out of band.
ENV_PAIRS=()
for var in GEMINI_MODEL REASONING_EFFORT DTR_CROSSCHECK IMAGE_PREPROCESS TIME_VERIFY \
           OCR_TIMEOUT VERIFY_TIMEOUT LOSSLESS_MIN_SIDE LOSSLESS_MAX_PIXELS; do
  if [[ -n "${!var:-}" ]]; then
    ENV_PAIRS+=("$var=${!var}")
  fi
done

ARGS=(
  run deploy "$SERVICE"
  --source "$STAGE"
  --region "$REGION"
  --allow-unauthenticated
  --memory "$MEMORY"
  --cpu "$CPU"
  --concurrency "$CONCURRENCY"
  --min-instances "$MIN_INSTANCES"
  --max-instances "$MAX_INSTANCES"
  --timeout "$TIMEOUT"
)
if ((${#ENV_PAIRS[@]})); then
  ARGS+=(--update-env-vars "$(IFS=,; echo "${ENV_PAIRS[*]}")")
fi
# Likewise --update-secrets, so a redeploy without GEMINI_SECRET leaves the
# already-attached key alone instead of detaching it.
if [[ -n "${GEMINI_SECRET:-}" ]]; then
  ARGS+=(--update-secrets "GEMINI_API_KEY=${GEMINI_SECRET}:latest")
fi

if [[ -n "${DRY_RUN:-}" ]]; then
  echo "DRY_RUN — would run:"
  printf '  gcloud'; printf ' %q' "${ARGS[@]}"; echo
  exit 0
fi

gcloud "${ARGS[@]}"

URL="$(gcloud run services describe "$SERVICE" --region "$REGION" \
        --format='value(status.url)' 2>/dev/null || true)"
if [[ -n "$URL" ]]; then
  echo
  echo "Service: $URL"
  echo -n "Health:  "
  curl -fsS --max-time 30 "$URL/api/health" || echo "(health check did not respond)"
  echo
fi
