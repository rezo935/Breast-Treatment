/*
========================================================================================
ESAPI C# SCRIPT: Breast Planning - Structure Preparation Workflow
Context: Varian Eclipse TPS.
Inputs needed: StructureSet, Image.
Goal: Create and modify structures for breast IMRT/VMAT planning using high-resolution segments.
========================================================================================
*/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Media.Media3D;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

/// <summary>
/// Provides methods to create and configure breast planning structures
/// in Varian Eclipse via the Eclipse Scripting API (ESAPI).
/// All structures are created as high-resolution segments.
/// </summary>
public class BreastPlanningStructures
{
    /// <summary>
    /// Creates all breast planning structures sequentially (Steps 1–5).
    /// </summary>
    /// <param name="structureSet">The active structure set to modify.</param>
    /// <param name="image">The CT image used for HU-based thresholding.</param>
    public void CreateBreastStructures(StructureSet structureSet, Image image)
    {
        // STEP 1: Standard Body
        CreateStandardBody(structureSet, image);

        // STEP 2: Actual Body
        Structure actualBody = CreateActualBody(structureSet, image);

        // STEP 3: Breathing Margin
        Structure breathingMargin = CreateBreathingMargin(structureSet);

        // STEP 4: Breathing Margin Crop
        CreateBreathingMarginCrop(structureSet, actualBody, breathingMargin);

        // STEP 5: Optimization structures
        CreateOptimizationStructures(structureSet, actualBody);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 1 – Standard Body
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Checks whether the "\body" structure already exists; if not, creates it
    /// using a –700 HU threshold on the CT image.
    /// </summary>
    private void CreateStandardBody(StructureSet structureSet, Image image)
    {
        Structure body = structureSet.Structures
            .FirstOrDefault(s => s.Id.Equals("\\body", StringComparison.OrdinalIgnoreCase));

        if (body == null)
        {
            body = structureSet.AddStructure("BODY", "\\body");
            body.ConvertToHighResolution();
            GenerateStructureFromHUThreshold(body, image, lowerHU: -700.0);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 2 – Actual Body
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Creates the "Actual Body" CONTROL structure using a –300 HU threshold.
    /// Colour: Magenta.
    /// </summary>
    private Structure CreateActualBody(StructureSet structureSet, Image image)
    {
        Structure actualBody = structureSet.AddStructure("CONTROL", "Actual Body");
        actualBody.Color = Color.Magenta;
        actualBody.ConvertToHighResolution();
        GenerateStructureFromHUThreshold(actualBody, image, lowerHU: -300.0);
        return actualBody;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 3 – Breathing Margin
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Creates the "Breathing Margin" PTV structure by applying an asymmetric
    /// outer margin to "CTV Breast". The lateral/medial margin sides are
    /// determined by the laterality of "CTV Breast".
    /// Colour: Orange.
    /// </summary>
    private Structure CreateBreathingMargin(StructureSet structureSet)
    {
        Structure ctvBreast = structureSet.Structures
            .FirstOrDefault(s => s.Id.Equals("CTV Breast", StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException(
                "Structure 'CTV Breast' not found in the structure set.");

        bool isLeftSided = IsLeftSided(ctvBreast);

        // All margins expressed in millimetres (ESAPI uses mm internally).
        //   2.0 cm = 20 mm  |  0.5 cm =  5 mm  |  0.8 cm = 8 mm
        //   1.0 cm = 10 mm  |  0.4 cm =  4 mm
        const double anteriorMargin  = 20.0; // 2.0 cm anterior
        const double posteriorMargin =  5.0; // 0.5 cm posterior
        const double supInfMargin    =  8.0; // 0.8 cm superior and inferior
        const double lateralMargin   = 10.0; // 1.0 cm towards the breast side
        const double medialMargin    =  4.0; // 0.4 cm towards the midline

        // Eclipse/DICOM patient coordinate axes for HFS:
        //   +X = patient LEFT   –X = patient RIGHT
        //   +Y = POSTERIOR      –Y = ANTERIOR
        //   +Z = SUPERIOR       –Z = INFERIOR
        //
        // AxisAlignedMargins(geometry, x1(–X/Right), x2(+X/Left),
        //                              y1(–Y/Ant),   y2(+Y/Post),
        //                              z1(–Z/Inf),   z2(+Z/Sup))
        double rightMargin = isLeftSided ? medialMargin  : lateralMargin;
        double leftMargin  = isLeftSided ? lateralMargin : medialMargin;

        var margins = new AxisAlignedMargins(
            StructureMarginGeometry.Outer,
            rightMargin, leftMargin,         // –X / +X
            anteriorMargin, posteriorMargin, // –Y / +Y
            supInfMargin, supInfMargin);     // –Z / +Z

        Structure breathingMargin = structureSet.AddStructure("PTV", "Breathing Margin");
        breathingMargin.Color = Color.Orange;
        breathingMargin.ConvertToHighResolution();
        breathingMargin.SegmentVolume = ctvBreast.SegmentVolume.AsymmetricMargin(margins);
        return breathingMargin;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 4 – Breathing Margin Crop
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Creates "Breathing Margin Crop" as the intersection of "Breathing Margin"
    /// with "Actual Body" contracted by 4 mm (0.4 cm inner margin).
    /// Colour: Orange.
    /// </summary>
    private void CreateBreathingMarginCrop(
        StructureSet structureSet,
        Structure actualBody,
        Structure breathingMargin)
    {
        // Contract Actual Body inward by 4 mm (= 0.4 cm) to define the crop boundary.
        // A negative argument to Margin() produces an inner (contracting) offset.
        SegmentVolume contractedBody = actualBody.SegmentVolume.Margin(-4.0);

        // Intersect with Breathing Margin
        SegmentVolume cropVolume = breathingMargin.SegmentVolume.And(contractedBody);

        Structure breathingMarginCrop = structureSet.AddStructure("PTV", "Breathing Margin Crop");
        breathingMarginCrop.Color = Color.Orange;
        breathingMarginCrop.ConvertToHighResolution();
        breathingMarginCrop.SegmentVolume = cropVolume;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 5 – Optimization Structures
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Duplicates CTV/PTV structures into "Opt" variants, then:
    ///  • CTV Opt: crops 4 mm inside Actual Body and subtracts both lungs.
    ///  • PTV Opt: crops 3 mm inside Actual Body and subtracts both lungs.
    /// Source structures that do not exist in the structure set are skipped.
    /// </summary>
    private void CreateOptimizationStructures(StructureSet structureSet, Structure actualBody)
    {
        Structure lungL = GetRequiredStructure(structureSet, "Lung L");
        Structure lungR = GetRequiredStructure(structureSet, "Lung R");

        // CTV Opt structures – crop 4 mm inside Actual Body
        DuplicateAndCropOpt(structureSet, "CTV Breast", "CTV Breast Opt",
            actualBody, lungL, lungR, cropMm: 4.0, dicomType: "CTV");

        DuplicateAndCropOpt(structureSet, "CTV LN", "CTV LN Opt",
            actualBody, lungL, lungR, cropMm: 4.0, dicomType: "CTV");

        DuplicateAndCropOpt(structureSet, "CTV IMN", "CTV IMN Opt",
            actualBody, lungL, lungR, cropMm: 4.0, dicomType: "CTV");

        // PTV Opt structures – crop 3 mm inside Actual Body
        DuplicateAndCropOpt(structureSet, "PTV LN", "PTV LN Opt",
            actualBody, lungL, lungR, cropMm: 3.0, dicomType: "PTV");

        DuplicateAndCropOpt(structureSet, "PTV IMN", "PTV IMN Opt",
            actualBody, lungL, lungR, cropMm: 3.0, dicomType: "PTV");
    }

    /// <summary>
    /// Creates a new "Opt" structure by duplicating <paramref name="sourceId"/>,
    /// cropping it to the contracted body, and subtracting both lung structures.
    /// </summary>
    private void DuplicateAndCropOpt(
        StructureSet structureSet,
        string sourceId,
        string targetId,
        Structure actualBody,
        Structure lungL,
        Structure lungR,
        double cropMm,
        string dicomType)
    {
        Structure source = structureSet.Structures
            .FirstOrDefault(s => s.Id.Equals(sourceId, StringComparison.OrdinalIgnoreCase));

        if (source == null)
            return; // Source structure absent – skip silently

        // Contract Actual Body inward by cropMm to define the crop boundary
        SegmentVolume croppedBody = actualBody.SegmentVolume.Margin(-cropMm);

        // Intersect source with the contracted body, then subtract lungs
        SegmentVolume vol = source.SegmentVolume
            .And(croppedBody)
            .Sub(lungL.SegmentVolume)
            .Sub(lungR.SegmentVolume);

        Structure target = structureSet.AddStructure(dicomType, targetId);
        target.ConvertToHighResolution();
        target.SegmentVolume = vol;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // HU-Threshold Contouring Helpers
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Populates <paramref name="structure"/> with contours derived by thresholding
    /// each CT slice at [<paramref name="lowerHU"/>, <paramref name="upperHU"/>].
    ///
    /// <para>
    /// ESAPI's <see cref="Image.GetVoxels"/> fills a buffer with the calibrated
    /// Hounsfield-unit value for every voxel on the requested image plane,
    /// so no additional rescale step is required.
    /// </para>
    /// <para>
    /// <see cref="Structure.ConvertToHighResolution"/> must be called on
    /// <paramref name="structure"/> before invoking this method.
    /// </para>
    /// </summary>
    private static void GenerateStructureFromHUThreshold(
        Structure structure,
        Image image,
        double lowerHU,
        double upperHU = 3071.0)
    {
        for (int z = 0; z < image.ZSize; z++)
        {
            int[,] voxels = new int[image.XSize, image.YSize];
            image.GetVoxels(z, voxels);

            foreach (VVector[] contour in ExtractContoursFromMask(voxels, image, z, lowerHU, upperHU))
            {
                if (contour.Length >= 3)
                    structure.AddContourOnImagePlane(contour, z);
            }
        }
    }

    /// <summary>
    /// Extracts a single closed-boundary polygon from one CT slice using a
    /// column-scanline approach:
    /// <list type="bullet">
    ///   <item>Left-to-right pass collects the topmost (minimum-row) threshold pixel per column.</item>
    ///   <item>Right-to-left pass collects the bottommost (maximum-row) threshold pixel per column.</item>
    /// </list>
    /// Columns that contain no threshold-passing voxels contribute no point to the polygon
    /// and are skipped; the contour polygon still closes naturally because adjacent contributing
    /// columns are connected in order.  This approach works correctly for the single,
    /// roughly convex cross-section produced by body-level HU thresholds (–700 HU / –300 HU).
    /// For slices with no threshold-passing voxels the enumeration yields nothing.
    /// </summary>
    private static IEnumerable<VVector[]> ExtractContoursFromMask(
        int[,] voxels,
        Image image,
        int z,
        double lowerHU,
        double upperHU)
    {
        int xSize = image.XSize;
        int ySize = image.YSize;

        // Build binary mask
        bool[,] mask = new bool[xSize, ySize];
        bool anySet = false;
        for (int x = 0; x < xSize; x++)
        {
            for (int y = 0; y < ySize; y++)
            {
                double hu = voxels[x, y];
                if (hu >= lowerHU && hu <= upperHU)
                {
                    mask[x, y] = true;
                    anySet = true;
                }
            }
        }

        if (!anySet)
            yield break;

        // Forward pass (left→right): for each column pick the topmost row that passes
        // the threshold.  Columns with no qualifying voxels are skipped (no point added).
        var points = new List<VVector>();
        for (int x = 0; x < xSize; x++)
        {
            for (int y = 0; y < ySize; y++)
            {
                if (mask[x, y])
                {
                    points.Add(PixelToPatient(image, x, y, z));
                    break; // stop at first (topmost) qualifying row in this column
                }
            }
        }

        // Reverse pass (right→left): for each column pick the bottommost row that passes
        // the threshold.  Columns with no qualifying voxels are skipped.
        // Together with the forward pass this traces the full outer boundary of the
        // thresholded region and forms a closed polygon.
        for (int x = xSize - 1; x >= 0; x--)
        {
            for (int y = ySize - 1; y >= 0; y--)
            {
                if (mask[x, y])
                {
                    points.Add(PixelToPatient(image, x, y, z));
                    break; // stop at first (bottommost) qualifying row in this column
                }
            }
        }

        if (points.Count >= 3)
            yield return points.ToArray();
    }

    /// <summary>
    /// Converts a voxel position (column <paramref name="x"/>, row <paramref name="y"/>,
    /// plane <paramref name="z"/>) into patient-coordinate millimetres using the
    /// image's geometric parameters.
    /// </summary>
    private static VVector PixelToPatient(Image image, int x, int y, int z)
    {
        // Patient position = Origin + x·XRes·RowDirection
        //                           + y·YRes·ColumnDirection
        //                           + z·ZRes·NormalDirection
        return new VVector(
            image.Origin.x
                + x * image.XRes * image.RowDirection.x
                + y * image.YRes * image.ColumnDirection.x
                + z * image.ZRes * image.NormalDirection.x,
            image.Origin.y
                + x * image.XRes * image.RowDirection.y
                + y * image.YRes * image.ColumnDirection.y
                + z * image.ZRes * image.NormalDirection.y,
            image.Origin.z
                + x * image.XRes * image.RowDirection.z
                + y * image.YRes * image.ColumnDirection.z
                + z * image.ZRes * image.NormalDirection.z);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Utility Helpers
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Returns <see langword="true"/> when the centroid of
    /// <paramref name="structure"/>'s bounding box lies in the positive-X half of
    /// patient space, which corresponds to the patient's LEFT side in standard
    /// Head-First Supine (HFS) orientation.
    /// </summary>
    private static bool IsLeftSided(Structure structure)
    {
        Rect3D bounds = structure.MeshGeometry.Bounds;
        double centerX = bounds.X + bounds.SizeX / 2.0;
        return centerX > 0.0;
    }

    /// <summary>
    /// Finds a structure by ID (case-insensitive) and throws a descriptive
    /// <see cref="InvalidOperationException"/> when it is absent.
    /// </summary>
    private static Structure GetRequiredStructure(StructureSet structureSet, string id)
    {
        return structureSet.Structures
            .FirstOrDefault(s => s.Id.Equals(id, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException(
                $"Required structure '{id}' not found in the structure set.");
    }
}
