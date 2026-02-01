# 2D-Object-Transformation
This project was created during the first semester of my second year in Creative Computing at South East Technological University.
# Assignment 01 — 2D Object Transformation

**Author:** Annetta Hawkins

---

## Overview

This project implements **2D geometric transformations** (rotation and translation) on a composite object using **homogeneous coordinates** and **transformation matrices** in an Excel macro-enabled workbook (`.xlsm`).

The object modeled is an airplane, composed of five distinct polygonal parts — Tail, Fuselage, Nose, Wing1, and Wing2. Each part is defined by a set of 2D vertices stored in homogeneous form (x, y, 1), and a 3×3 transformation matrix is applied to produce rotated/translated coordinates.

---

## File Structure

| File | Description |
|---|---|
| `Annetta_Hawkins_Assignment01_Submission.xlsm` | Main macro-enabled Excel workbook containing all object data, the user interface, and transformation logic. |

---

## Workbook Sheets

### `Object`
Stores the **original and transformed vertex data** for the airplane and its sub-components.

- **Top section:** Defines the object's bounding box dimensions (Length = 330, Width = 280) and the centre of the object (`cx`, `cy`).
- **Per-part sections:** Each airplane component (Tail, Fuselage, Nose, Wing1, Wing2) is defined by:
  - A label column listing the vertex identifiers (e.g., `a`, `b`, `c`, …).
  - Original x-coordinates, y-coordinates, and homogeneous w-coordinates (all set to 1).
  - Transformed x′, y′, and w′ columns produced by multiplying the transformation matrix by each vertex.

### `Interface`
Provides the **user-facing controls** for the transformation.

| Control | Description |
|---|---|
| **Centre of Rotation (`cx`, `cy`)** | The pivot point around which the object rotates. Defaults to (0, 0). |
| **Angle** | The rotation angle in degrees (currently set to ≈ 23.88°). |
| **Spin** | A secondary rotation or continuous rotation parameter. |

Changing these values triggers a recalculation of the transformation matrix and all transformed vertex coordinates.

### `Transform`
Contains the **3×3 homogeneous transformation matrices** used to compute the new vertex positions.

The matrices follow the standard 2D rotation + translation pipeline:

1. **Translate** the object so that the centre of rotation moves to the origin.
2. **Rotate** by the specified angle using the rotation matrix:

```
[ cos θ   −sin θ   0 ]
[ sin θ    cos θ   0 ]
[   0        0     1 ]
```

3. **Translate back** to the original centre of rotation.

The sheet also displays the intermediate and final combined matrices for verification.

---

## How It Works

1. The user sets a **centre of rotation** and a **rotation angle** on the `Interface` sheet.
2. A combined **translation + rotation** matrix is assembled on the `Transform` sheet using homogeneous coordinates.
3. For every vertex of every airplane part on the `Object` sheet, the transformation matrix is multiplied by the vertex vector `[x, y, 1]ᵀ` to produce the new coordinates `[x′, y′, 1]ᵀ`.
4. The transformed coordinates are stored alongside the originals, ready for plotting or further use.

---

## Key Concepts

- **Homogeneous Coordinates:** Each 2D point `(x, y)` is represented as `(x, y, 1)`, enabling translations to be expressed as matrix multiplications.
- **Rotation Matrix:** A 3×3 matrix that rotates all points by angle θ around a chosen pivot.
- **Composite Object:** The airplane is split into five independently defined polygons (Tail, Fuselage, Nose, Wing1, Wing2), each transformed using the same global matrix.

---

## How to Use

1. Open `Annetta_Hawkins_Assignment01_Submission.xlsm` in **Microsoft Excel** (macros must be enabled).
2. Navigate to the **`Interface`** sheet.
3. Adjust the **Centre of Rotation** (`cx`, `cy`) and the **Angle** (in degrees) as needed.
4. The transformed coordinates on the **`Object`** sheet and the matrices on the **`Transform`** sheet will update automatically.

> **Note:** Because the workbook uses `.xlsm` format, you may be prompted to enable macros when opening. Allow macros for all features to function correctly.

---

## Requirements

- Microsoft Excel (desktop version recommended; some features may not work in Excel Online or Google Sheets).
- Macros enabled.
