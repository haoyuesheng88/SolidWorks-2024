---
name: solidworks-geometry-airflow
description: Connect to an already open SOLIDWORKS session on Windows, inspect the active part, read selected face area, and compute airflow and velocity pressure from the selected inlet face. Use when the user asks to connect to SOLIDWORKS, read a live model, count bodies or faces, calculate volume or selected face area, or estimate duct airflow from face area and air velocity.
---

# SOLIDWORKS Geometry Airflow

Use this skill when SOLIDWORKS is already open and the user wants measurements from the live model instead of estimating from a screenshot.

Assume Windows, PowerShell, Python 3, and `pywin32` are available. Keep the workflow read-only unless the user explicitly asks to modify the model.

## Workflow

1. Confirm that `SLDWORKS` is running.
2. Connect to the active SOLIDWORKS application through COM.
3. Read the active document and verify it is a part document.
4. Run the appropriate bundled script for the requested metric.
5. Report the measured values with units and note any assumptions used in airflow calculations.

## Bundled Scripts

- `scripts/solidworks_active_part_metrics.py`
  Read the active part and output title, body count, face count, volume, and surface area.
- `scripts/solidworks_selected_face_metrics.py`
  Read the first selected face in the active part and output its area.
- `scripts/solidworks_selected_face_airflow.py`
  Read the first selected face, apply a supplied air velocity, and output area, volumetric flow, and velocity pressure.

## Default Commands

Read active part metrics:

```powershell
python ".\skills\solidworks-geometry-airflow\scripts\solidworks_active_part_metrics.py"
```

Read selected face area:

```powershell
python ".\skills\solidworks-geometry-airflow\scripts\solidworks_selected_face_metrics.py"
```

Compute airflow from the selected face at `12 m/s`:

```powershell
python ".\skills\solidworks-geometry-airflow\scripts\solidworks_selected_face_airflow.py" --velocity 12
```

## Important Rules

- Treat the model as the source of truth. Prefer live SOLIDWORKS measurements over screenshot estimates.
- Require an already open SOLIDWORKS session. Do not launch SOLIDWORKS unless the user explicitly asks.
- Require an active part document for these scripts. If the active document is an assembly or drawing, report that clearly.
- For face-based calculations, use the first selected face and tell the user if nothing is selected.
- Report airflow in `m^3/s`, `L/s`, and `m^3/h`.
- Report velocity pressure as `0.5 * density * velocity^2`.
- Do not present velocity pressure as system static pressure. Static pressure cannot be uniquely determined from face area and velocity alone.

## Interpretation Rules

- `BODY_COUNT` is the number of visible solid bodies returned by `GetBodies2(0, True)`.
- `FACE_COUNT` is the sum of `GetFaceCount()` across the returned solid bodies.
- Volume comes from the active part mass properties and is reported in both `m^3` and `mm^3`.
- Selected face area comes from `IFace2.GetArea()` and is reported in both `m^2` and `mm^2`.
- Airflow uses `Q = A * v`.
- Velocity pressure uses air density `1.2 kg/m^3` by default unless the user provides another density.

## What To Report

Include the requested metrics, units, and any assumptions.

For airflow questions, include:

- selected face area
- velocity used
- volumetric flow in `m^3/s`, `L/s`, and `m^3/h`
- velocity pressure in `Pa`
- a note that static pressure is not determined from velocity and area alone
