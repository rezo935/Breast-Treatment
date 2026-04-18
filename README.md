# Breast-Treatment
ESAPI base automation for breast treatments

## BreastLRv3 flexible beam-angle generation

A new single-file script is available at:
`/home/runner/work/Breast-Treatment/Breast-Treatment/BreastLRv3/BreastLRv3.cs`

### What it does
- WPF dialog captures:
  - Medial0 gantry (required)
  - collimator / couch (optional)
  - laterality (Left/Right)
  - LN/IMN inclusion (9-field vs 12-field template)
  - whether Medial0 beam already exists (create remaining beams only)
- Previews calculated angles before beam creation.
- Inserts beams into the active external beam plan using a non-setup beam as template.

### Offset definition (modifiable)
In `BreastBeamAngleCalculator`:
- `OffsetsBreastOnly` defines the 9-field offsets.
- `OffsetsWithLnImn` defines the 12-field offsets.

Offsets are degrees relative to Medial0.
- Left breast: offset is applied directly (`Medial0 + offset`).
- Right breast: offset is mirrored (`Medial0 - offset`).

If you need a different institutional beam pattern, modify those two arrays only.
