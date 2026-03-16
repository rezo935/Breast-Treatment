/*
========================================================================================
ESAPI C# SCRIPT: Smart, Modular Breast Structure Preparation
Context: Varian Eclipse TPS.
Goal: Automatically detect laterality and the presence of optional targets (Boost, LN, IMN)
to dynamically execute only the necessary structure creation steps.
========================================================================================
*/

using System;
using System.Collections.Generic;
using System.Windows.Media;
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
    /// Smart, modular entry point.  Detects laterality and the presence of
    /// optional targets (Boost, LN, IMN) and then executes only the steps
    /// that are required for the current patient.
    /// </summary>
    /// <param name="structureSet">The active structure set to modify.</param>
    /// <param name="image">The CT image used for HU-based thresholding.</param>
    public void CreateBreastStructures(StructureSet structureSet, Image image)
    {
        // ── STEP 1: Automatic Detection & Setup ──────────────────────────────

        // "CTV Breast" is mandatory; abort with a clear message if it is absent.
        Structure ctvBreast = structureSet.Structures
            .FirstOrDefault(s => s.Id.Equals("CTV Breast", StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException(
                "Structure 'CTV Breast' not found in the structure set.");

        // Laterality is determined from the structure ID (not geometry) so that
        // the user's naming convention ("Left", "L ", "L_") is respected exactly.
        bool isLeft = DetectLaterality(ctvBreast.Id);

        // Optional-module flags drive which later steps are executed.
        bool hasBoost = StructureExists(structureSet, "PTV Boost");
        bool hasLN    = StructureExists(structureSet, "CTV LN") ||
                        StructureExists(structureSet, "PTV LN");
        bool hasIMN   = StructureExists(structureSet, "CTV IMN") ||
                        StructureExists(structureSet, "PTV IMN");

        // Pre-fetch lung structures once; both are required for all crop operations.
        Structure lungL = GetRequiredStructure(structureSet, "Lung L");
        Structure lungR = GetRequiredStructure(structureSet, "Lung R");

        // ── STEP 2: Base Structures (always executed) ─────────────────────────

        CreateStandardBody(structureSet, image);

        Structure actualBody = CreateActualBody(structureSet, image);

        Structure breathingMargin = CreateBreathingMargin(structureSet, ctvBreast, isLeft);

        Structure breathingMarginCrop = CreateBreathingMarginCrop(structureSet, actualBody, breathingMargin);

        CreateCtvBreastOpt(structureSet, ctvBreast, actualBody, lungL, lungR);

        // ── STEP 3: Nodal Structures (only when nodal targets are present) ────

        Structure ptvLnOpt  = null;
        Structure ptvImnOpt = null;

        if (hasLN || hasIMN)
            CreateNodalStructures(structureSet, actualBody, lungL, lungR,
                hasLN, hasIMN, out ptvLnOpt, out ptvImnOpt);

        // ── STEP 4: SIB Logic (only when a boost target is present) ───────────

        if (hasBoost)
            CreateSibStructures(structureSet, breathingMarginCrop, ptvLnOpt, ptvImnOpt);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 1 Helpers – Detection
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Returns <see langword="true"/> when <paramref name="structureId"/> indicates
    /// a left-sided breast target.  Matches "Left" (case-insensitive), "L " or "L_".
    /// If none of the left-side tokens are found, right side is assumed.
    /// </summary>
    private static bool DetectLaterality(string structureId)
    {
        return structureId.IndexOf("Left", StringComparison.OrdinalIgnoreCase) >= 0
            || structureId.IndexOf("L ", StringComparison.OrdinalIgnoreCase) >= 0
            || structureId.IndexOf("L_", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    /// <summary>
    /// Returns <see langword="true"/> when a structure with <paramref name="id"/>
    /// (case-insensitive) is present in <paramref name="structureSet"/>.
    /// </summary>
    private static bool StructureExists(StructureSet structureSet, string id)
    {
        return structureSet.Structures
            .Any(s => s.Id.Equals(id, StringComparison.OrdinalIgnoreCase));
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 2 – Standard Body
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
        actualBody.Color = Colors.Magenta;
        actualBody.ConvertToHighResolution();
        GenerateStructureFromHUThreshold(actualBody, image, lowerHU: -300.0);
        return actualBody;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 2 – Breathing Margin
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Creates the "Breathing Margin" PTV structure by applying an asymmetric
    /// outer margin to <paramref name="ctvBreast"/>.  Laterality is supplied via
    /// <paramref name="isLeft"/> (detected from the structure ID in STEP 1).
    /// <list type="bullet">
    ///   <item>Left breast : Left 1.0 cm / Right 0.4 cm</item>
    ///   <item>Right breast: Left 0.4 cm / Right 1.0 cm</item>
    ///   <item>Anterior 2.0 cm, Posterior 0.5 cm, Superior/Inferior 0.8 cm (both sides)</item>
    /// </list>
    /// Colour: Orange.
    /// </summary>
    private static Structure CreateBreathingMargin(
        StructureSet structureSet,
        Structure ctvBreast,
        bool isLeft)
    {
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
        double rightMargin = isLeft ? medialMargin  : lateralMargin;
        double leftMargin  = isLeft ? lateralMargin : medialMargin;

        var margins = new AxisAlignedMargins(
            StructureMarginGeometry.Outer,
            rightMargin, leftMargin,         // –X / +X
            anteriorMargin, posteriorMargin, // –Y / +Y
            supInfMargin, supInfMargin);     // –Z / +Z

        Structure breathingMargin = structureSet.AddStructure("PTV", "Breathing Margin");
        breathingMargin.Color = Colors.Orange;
        breathingMargin.ConvertToHighResolution();
        breathingMargin.SegmentVolume = ctvBreast.SegmentVolume.AsymmetricMargin(margins);
        return breathingMargin;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 2 – Breathing Margin Crop
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Creates "Breathing Margin Crop" as the intersection of
    /// <paramref name="breathingMargin"/> with "Actual Body" contracted by 4 mm
    /// (0.4 cm inner margin).
    /// Colour: Orange.
    /// </summary>
    private static Structure CreateBreathingMarginCrop(
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
        breathingMarginCrop.Color = Colors.Orange;
        breathingMarginCrop.ConvertToHighResolution();
        breathingMarginCrop.SegmentVolume = cropVolume;
        return breathingMarginCrop;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 2 – CTV Breast Opt
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Creates "CTV Breast Opt" by cropping <paramref name="ctvBreast"/> 4 mm
    /// inside <paramref name="actualBody"/> and subtracting both lung structures.
    /// </summary>
    private static void CreateCtvBreastOpt(
        StructureSet structureSet,
        Structure ctvBreast,
        Structure actualBody,
        Structure lungL,
        Structure lungR)
    {
        SegmentVolume vol = ctvBreast.SegmentVolume
            .And(actualBody.SegmentVolume.Margin(-4.0))
            .Sub(lungL.SegmentVolume)
            .Sub(lungR.SegmentVolume);

        Structure opt = structureSet.AddStructure("CTV", "CTV Breast Opt");
        opt.ConvertToHighResolution();
        opt.SegmentVolume = vol;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 3 – Nodal Structures
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Creates optimisation variants of nodal CTV/PTV structures.
    /// <list type="bullet">
    ///   <item>CTV Opt: cropped 4 mm inside Actual Body, lungs subtracted.</item>
    ///   <item>PTV Opt: cropped 3 mm inside Actual Body, lungs subtracted.</item>
    /// </list>
    /// Source structures that are absent in the structure set are skipped silently.
    /// The created "PTV LN Opt" and "PTV IMN Opt" structures are returned via
    /// <paramref name="ptvLnOpt"/> and <paramref name="ptvImnOpt"/> so that STEP 4
    /// can incorporate them into PTV Total.
    /// </summary>
    private static void CreateNodalStructures(
        StructureSet structureSet,
        Structure actualBody,
        Structure lungL,
        Structure lungR,
        bool hasLN,
        bool hasIMN,
        out Structure ptvLnOpt,
        out Structure ptvImnOpt)
    {
        ptvLnOpt  = null;
        ptvImnOpt = null;

        if (hasLN)
        {
            DuplicateAndCropOpt(structureSet, "CTV LN", "CTV LN Opt",
                actualBody, lungL, lungR, cropMm: 4.0, dicomType: "CTV");

            ptvLnOpt = DuplicateAndCropOpt(structureSet, "PTV LN", "PTV LN Opt",
                actualBody, lungL, lungR, cropMm: 3.0, dicomType: "PTV");
        }

        if (hasIMN)
        {
            DuplicateAndCropOpt(structureSet, "CTV IMN", "CTV IMN Opt",
                actualBody, lungL, lungR, cropMm: 4.0, dicomType: "CTV");

            ptvImnOpt = DuplicateAndCropOpt(structureSet, "PTV IMN", "PTV IMN Opt",
                actualBody, lungL, lungR, cropMm: 3.0, dicomType: "PTV");
        }
    }

    /// <summary>
    /// Creates a new optimisation structure from <paramref name="sourceId"/> by
    /// intersecting it with <paramref name="actualBody"/> contracted inward by
    /// <paramref name="cropMm"/> millimetres, then subtracting both lungs.
    /// Returns the created <see cref="Structure"/>, or <see langword="null"/> when
    /// the source structure is not present in the structure set.
    /// </summary>
    private static Structure DuplicateAndCropOpt(
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
            return null; // Source structure absent – skip silently

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
        return target;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 4 – SIB Structures
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Creates the three SIB (Simultaneous Integrated Boost) planning structures:
    /// <list type="bullet">
    ///   <item>
    ///     <term>PTV Total</term>
    ///     <description>
    ///       Union of "Breathing Margin Crop" with any existing nodal PTV Opt
    ///       structures ("PTV LN Opt", "PTV IMN Opt").
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <term>PTV Elective</term>
    ///     <description>
    ///       "PTV Total" minus a 1.0 cm outer expansion of "PTV Boost".
    ///     </description>
    ///   </item>
    ///   <item>
    ///     <term>PTV TR</term>
    ///     <description>
    ///       "PTV Total" minus "PTV Boost" minus "PTV Elective".
    ///     </description>
    ///   </item>
    /// </list>
    /// </summary>
    private static void CreateSibStructures(
        StructureSet structureSet,
        Structure breathingMarginCrop,
        Structure ptvLnOpt,
        Structure ptvImnOpt)
    {
        Structure ptvBoost = GetRequiredStructure(structureSet, "PTV Boost");

        // PTV Total = Breathing Margin Crop ∪ PTV LN Opt (opt.) ∪ PTV IMN Opt (opt.)
        SegmentVolume ptvTotalVol = breathingMarginCrop.SegmentVolume;
        if (ptvLnOpt != null)
            ptvTotalVol = ptvTotalVol.Or(ptvLnOpt.SegmentVolume);
        if (ptvImnOpt != null)
            ptvTotalVol = ptvTotalVol.Or(ptvImnOpt.SegmentVolume);

        Structure ptvTotal = structureSet.AddStructure("PTV", "PTV Total");
        ptvTotal.ConvertToHighResolution();
        ptvTotal.SegmentVolume = ptvTotalVol;

        // PTV Elective = PTV Total − (PTV Boost expanded 1.0 cm outward)
        SegmentVolume boostExpanded = ptvBoost.SegmentVolume.Margin(10.0); // 10 mm = 1.0 cm
        SegmentVolume ptvElectiveVol = ptvTotalVol.Sub(boostExpanded);

        Structure ptvElective = structureSet.AddStructure("PTV", "PTV Elective");
        ptvElective.ConvertToHighResolution();
        ptvElective.SegmentVolume = ptvElectiveVol;

        // PTV TR = PTV Total − PTV Boost − PTV Elective
        SegmentVolume ptvTrVol = ptvTotalVol
            .Sub(ptvBoost.SegmentVolume)
            .Sub(ptvElectiveVol);

        Structure ptvTr = structureSet.AddStructure("PTV", "PTV TR");
        ptvTr.ConvertToHighResolution();
        ptvTr.SegmentVolume = ptvTrVol;
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
        // Patient position = Origin + x·XRes·XDirection
        //                           + y·YRes·YDirection
        //                           + z·ZRes·ZDirection
        return new VVector(
            image.Origin.x
                + x * image.XRes * image.XDirection.x
                + y * image.YRes * image.YDirection.x
                + z * image.ZRes * image.ZDirection.x,
            image.Origin.y
                + x * image.XRes * image.XDirection.y
                + y * image.YRes * image.YDirection.y
                + z * image.ZRes * image.ZDirection.y,
            image.Origin.z
                + x * image.XRes * image.XDirection.z
                + y * image.YRes * image.YDirection.z
                + z * image.ZRes * image.ZDirection.z);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Utility Helpers
    // ─────────────────────────────────────────────────────────────────────────

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
