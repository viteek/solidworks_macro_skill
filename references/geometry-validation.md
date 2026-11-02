# Geometry validation of exported STEP / DXF (optional)

Read this when the user wants exported files *checked* before they go to a customer, laser or CAM -
not for ordinary export macros. Requires Windows, Python 3.12 and `requirements-geometry.txt`
(ezdxf, shapely, cadquery-ocp, vtk). No CAM/machine is configured: every report says
`cam_status=NOT_CONFIGURED`; PASS means the declared geometric checks passed, nothing more.

## Contents
1. What the validators check (and don't)
2. Validated export pipeline
3. Checking existing files
4. DXF profiles - SOLIDWORKS specifics (measured)
5. STEP tolerances
6. Acceptance once a CAM is chosen

## 1. What the validators check

| Tool | Checks | Limit |
|---|---|---|
| `validate_step.py` | STEP parsed by OpenCascade (independent kernel); explicit length unit in the geometric context (no unit = rejected, never assumed mm); valid closed solids; no loose faces/edges; per-body volume, area, centroid, bounds vs SOLIDWORKS source metrics of the **same configuration** | Metric agreement is not a full shape-equivalence proof |
| `validate_dxf.py` | Declared units, operation role per entity (linetype, then layer), closed cut contours, holes/net area, self-intersections, overlapping or T-junction cut paths, XY plane, bend count | Screening, not comparison with an approved flat pattern; material model approximate for formed parts (see §4) |
| `validated_export.py` | Export to a private job folder → validation → publication of the validated bytes only (hash-checked); existing targets are never replaced, including with `--overwrite` | Parts → STEP/DXF only; assemblies and PDF are not geometry-validated |

## 2. Validated export pipeline (runs SOLIDWORKS via `run_in_solidworks.py`)

```text
python scripts/validated_export.py C:\SWTest\BLACHA.sldprt C:\SWTest\out\BLACHA.dxf --profile profiles/solidworks-sheetmetal-mm.json --normalize-units --json C:\SWTest\out\dxf.json
python scripts/validated_export.py C:\SWTest\PART.sldprt C:\SWTest\out\PART.step --tolerances profiles/step-tolerances.json [--configuration L200] --json C:\SWTest\out\step.json
```
Statuses: `PUBLISHED` (file created), `BLOCKED_EXPORT` (macro/runner failed - see `runner`, `macro_log`),
`BLOCKED_VALIDATION` (see `validation.issues`; previous output untouched), `ERROR`. The job folder
`.sw-validation-*` is deleted after publication and kept otherwise for inspection.
Existing target files are rejected even when `--overwrite` is supplied; no `.previous-*` backup is created. Choose a new output path.
Use test copies of parts: `--open` loads them into the user's SOLIDWORKS.

## 3. Checking existing files (no SOLIDWORKS)

```text
python scripts/validate_dxf.py part.dxf --profile profiles/solidworks-sheetmetal-mm.json --json check.json
python scripts/validate_dxf.py part.dxf --profile profiles/solidworks-sheetmetal-mm.json --normalize-to part_mm.dxf --layer-map "{}" --json norm.json
python scripts/validate_step.py part.step --source-metrics source-metrics.json --tolerances profiles/step-tolerances.json --json check.json
```
`--normalize-to` writes a NEW mm copy (never overwrites); `--layer-map` renames layers explicitly.
Source metrics come from `assets/templates/source_metrics_template.bas` (visible solid bodies, active
configuration, mm).

## 4. DXF profiles - SOLIDWORKS specifics (measured on SW 2024 SP5)

- A flat-pattern DXF from `ExportToDWG2` (mask 1+4+8) has the outline AND the bend lines on layer
  **0**; bend lines differ only by linetype **CENTERX2**. Sketch points are written as `POINT`.
  Hence `profiles/solidworks-sheetmetal-mm.json` maps `layers {"0": "cut"}`, `linetypes {"CENTERX2": "bend", ...}`
  and ignores `POINT`. A layer-only mapping (`0 → cut`) would turn bend lines into cut paths.
- Units follow the part: inch parts give inch DXFs → `UNITS_MISMATCH` unless `--normalize-units`
  (exact factor 25.4, new copy).
- Results on SOLIDWORKS sample parts: simple sheet-metal part → PASS after normalization, 19 bend
  lines found; complex part → `CUT_INTERSECTION` x2 + `OPEN_CONTOUR` at a real T-junction of cut
  lines (a genuine CAM concern - review those handles in CAD).
- Bend/mark lines not fully inside the computed material are a **warning**: bend lines run across
  relief cut-outs by design and the ring-parity material model is approximate for louvers/forms.
- Tolerances: topology 0.001 mm, curve 0.01 mm. 1e-6 mm (the 2.1 default) produced false failures
  after unit conversion.
- To make the check meaningful for a specific part, add `"expected"` to a copy of the profile -
  e.g. `"bend_entities": 4` catches a DXF exported without bend lines (the mask-5 problem). Values
  must come from the drawing/approved flat pattern, never from the file being checked.
- `profiles/neutral-mm.json` (layers CUT/BEND/MARK) suits DXFs already organised by layer.

## 5. STEP tolerances

`profiles/step-tolerances.json`: linear 0.01 mm, area 0.01 mm² + 1e-4 relative, volume 0.01 mm³ + 1e-4
relative - data-exchange checks, not manufacturing tolerances. Tighten only with evidence; freeform
parts need the relative term.

## 6. Acceptance once a CAM is chosen

Record importer/machine/post-processor versions; import a known part 1:1 and measure it; map CUT/BEND/MARK
to real operations; confirm thickness, K-factor and bend directions with production; set kerf
compensation, lead-ins and pierces in CAM (the exporter doesn't); simulate and cut a sample; store the
result with the profile and the input file hash. Until then `cam_status` stays `NOT_CONFIGURED`.
