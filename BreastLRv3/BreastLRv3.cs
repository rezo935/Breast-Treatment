using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        public void Execute(ScriptContext context /*, Window window, ScriptEnvironment environment*/)
        {
            if (context == null || context.Patient == null || context.Course == null || context.ExternalPlanSetup == null)
            {
                MessageBox.Show("Open a patient, course, and external beam plan before running this script.", "BreastLRv3");
                return;
            }

            var plan = context.ExternalPlanSetup;
            var defaults = BuildDefaults(context, plan);

            var dialog = new BeamAngleDialog(defaults);
            var result = dialog.ShowDialog();
            if (!result.HasValue || !result.Value)
                return;

            var request = dialog.SelectedInput;
            var preview = dialog.PreviewFields;

            if (request == null || preview == null || preview.Count == 0)
            {
                MessageBox.Show("No calculated fields were provided.", "BreastLRv3");
                return;
            }

            var log = new StringBuilder();
            log.AppendLine("BreastLRv3 angle generation");
            log.AppendLine("---------------------------");
            log.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "Inputs: Medial0={0:0.##}°, Collimator={1:0.##}°, Couch={2:0.##}°, Laterality={3}, LN/IMN={4}, RemainingOnly={5}",
                request.Medial0Gantry,
                request.Collimator,
                request.Couch,
                request.Laterality == BreastLaterality.Left ? "Left" : "Right",
                request.IncludeLnImn ? "Yes" : "No",
                request.Medial0AlreadyExists ? "Yes" : "No"));

            try
            {
                context.Patient.BeginModifications();

                int created = InsertBeamsFromPreview(plan, request, preview, log);
                log.AppendLine(string.Format("Done. Created {0} beam(s).", created));
            }
            catch (Exception ex)
            {
                log.AppendLine("Error: " + ex.Message);
                MessageBox.Show(log.ToString(), "BreastLRv3");
                return;
            }

            MessageBox.Show(log.ToString(), "BreastLRv3");
        }

        private static int InsertBeamsFromPreview(
            ExternalPlanSetup plan,
            BeamAngleInput input,
            IList<BeamAngleField> preview,
            StringBuilder log)
        {
            var templateBeam = plan.Beams.FirstOrDefault(b => !b.IsSetupField);
            if (templateBeam == null)
                throw new InvalidOperationException("At least one treatment beam is required in the plan as a template.");

            if (templateBeam.ControlPoints == null || templateBeam.ControlPoints.Count() == 0)
                throw new InvalidOperationException("Template beam does not contain control points.");

            var cp = templateBeam.ControlPoints.First();
            var jaw = cp.JawPositions;
            var iso = templateBeam.IsocenterPosition;
            var machine = new ExternalBeamMachineParameters(
                templateBeam.TreatmentUnit.Id,
                templateBeam.EnergyModeDisplayName,
                templateBeam.DoseRate,
                "STATIC",
                null);

            int created = 0;
            for (int i = 0; i < preview.Count; i++)
            {
                var field = preview[i];

                if (input.Medial0AlreadyExists && field.OffsetFromMedial0 == 0.0)
                {
                    log.AppendLine("Skipping Medial0 (already exists). Angle: " + field.Gantry.ToString("0.##", CultureInfo.InvariantCulture) + "°");
                    continue;
                }

                try
                {
                    var beam = plan.AddStaticBeam(
                        machine,
                        jaw,
                        field.Collimator,
                        field.Gantry,
                        field.Couch,
                        iso);

                    beam.Id = BuildUniqueBeamId(plan, field.Name, i + 1);

                    created++;
                    log.AppendLine(string.Format(CultureInfo.InvariantCulture,
                        "Created {0}: Gantry {1:0.##}°, Collimator {2:0.##}°, Couch {3:0.##}°",
                        beam.Id,
                        field.Gantry,
                        field.Collimator,
                        field.Couch));
                }
                catch (Exception ex)
                {
                    log.AppendLine("Failed to create field " + field.Name + ": " + ex.Message);
                }
            }

            return created;
        }

        private static string BuildUniqueBeamId(ExternalPlanSetup plan, string baseName, int index)
        {
            string root = baseName;
            if (string.IsNullOrWhiteSpace(root))
                root = "BRT" + index.ToString("00", CultureInfo.InvariantCulture);

            root = root.Length > 12 ? root.Substring(0, 12) : root;
            string candidate = root;
            int suffix = 1;

            while (plan.Beams.Any(b => string.Equals(b.Id, candidate, StringComparison.OrdinalIgnoreCase)))
            {
                string suffixText = suffix.ToString(CultureInfo.InvariantCulture);
                int maxBase = Math.Max(1, 16 - suffixText.Length);
                string trimmedRoot = root.Length > maxBase ? root.Substring(0, maxBase) : root;
                candidate = trimmedRoot + suffixText;
                suffix++;
            }

            return candidate;
        }

        private static BeamAngleInput BuildDefaults(ScriptContext context, ExternalPlanSetup plan)
        {
            var input = new BeamAngleInput();
            input.Collimator = 0.0;
            input.Couch = 0.0;
            input.Medial0Gantry = 300.0;
            input.Laterality = GuessLaterality(context);
            input.IncludeLnImn = false;
            input.Medial0AlreadyExists = false;

            var beam = plan.Beams.FirstOrDefault(b => !b.IsSetupField);
            if (beam != null)
            {
                input.Medial0Gantry = NormalizeAngle(beam.GantryAngle);
                input.Collimator = NormalizeAngle(beam.CollimatorAngle);
                input.Couch = NormalizeAngle(beam.PatientSupportAngle);
            }

            return input;
        }

        private static BreastLaterality GuessLaterality(ScriptContext context)
        {
            if (context.StructureSet == null)
                return BreastLaterality.Left;

            var ids = context.StructureSet.Structures.Select(s => s.Id).ToList();
            bool left = ids.Any(id =>
                id.IndexOf("left", StringComparison.OrdinalIgnoreCase) >= 0 ||
                id.IndexOf("_l", StringComparison.OrdinalIgnoreCase) >= 0 ||
                id.EndsWith(" l", StringComparison.OrdinalIgnoreCase));

            return left ? BreastLaterality.Left : BreastLaterality.Right;
        }

        private static double NormalizeAngle(double angle)
        {
            double normalized = angle % 360.0;
            if (normalized < 0.0)
                normalized += 360.0;
            return normalized;
        }
    }

    internal enum BreastLaterality
    {
        Left,
        Right
    }

    internal sealed class BeamAngleInput
    {
        public double Medial0Gantry { get; set; }
        public double Collimator { get; set; }
        public double Couch { get; set; }
        public BreastLaterality Laterality { get; set; }
        public bool IncludeLnImn { get; set; }
        public bool Medial0AlreadyExists { get; set; }
    }

    internal sealed class BeamAngleField
    {
        public string Name { get; set; }
        public double Gantry { get; set; }
        public double Collimator { get; set; }
        public double Couch { get; set; }
        public double OffsetFromMedial0 { get; set; }
    }

    internal static class BreastBeamAngleCalculator
    {
        // Offsets are in degrees relative to user-entered Medial0 gantry.
        // Left breast applies offsets directly. Right breast mirrors offsets.
        private static readonly double[] OffsetsBreastOnly = new[]
        {
            0.0, 4.0, -4.0, 8.0, -8.0, 12.0, -12.0, 16.0, -16.0
        };

        private static readonly double[] OffsetsWithLnImn = new[]
        {
            0.0, 4.0, -4.0, 8.0, -8.0, 12.0, -12.0, 16.0, -16.0, 20.0, -20.0, 24.0
        };

        public static IList<BeamAngleField> Calculate(BeamAngleInput input)
        {
            if (input == null)
                throw new ArgumentNullException("input");

            var offsets = input.IncludeLnImn ? OffsetsWithLnImn : OffsetsBreastOnly;
            var fields = new List<BeamAngleField>(offsets.Length);

            for (int i = 0; i < offsets.Length; i++)
            {
                double signedOffset = input.Laterality == BreastLaterality.Left ? offsets[i] : -offsets[i];
                double gantry = NormalizeAngle(input.Medial0Gantry + signedOffset);

                fields.Add(new BeamAngleField
                {
                    Name = "BRT" + (i + 1).ToString("00", CultureInfo.InvariantCulture),
                    Gantry = gantry,
                    Collimator = NormalizeAngle(input.Collimator),
                    Couch = NormalizeAngle(input.Couch),
                    OffsetFromMedial0 = signedOffset
                });
            }

            return fields;
        }

        public static bool TryValidateInput(
            string medial0Text,
            string collimatorText,
            string couchText,
            BreastLaterality laterality,
            bool includeLnImn,
            bool medial0Exists,
            out BeamAngleInput input,
            out string error)
        {
            input = null;
            error = null;

            double medial0;
            if (!TryParseAngle(medial0Text, out medial0))
            {
                error = "Medial0 gantry angle is invalid. Enter a number in [0, 360].";
                return false;
            }

            double collimator;
            if (string.IsNullOrWhiteSpace(collimatorText))
            {
                collimator = 0.0;
            }
            else if (!TryParseAngle(collimatorText, out collimator))
            {
                error = "Collimator angle is invalid. Enter a number in [0, 360].";
                return false;
            }

            double couch;
            if (string.IsNullOrWhiteSpace(couchText))
            {
                couch = 0.0;
            }
            else if (!TryParseAngle(couchText, out couch))
            {
                error = "Couch angle is invalid. Enter a number in [0, 360].";
                return false;
            }

            input = new BeamAngleInput
            {
                Medial0Gantry = NormalizeAngle(medial0),
                Collimator = NormalizeAngle(collimator),
                Couch = NormalizeAngle(couch),
                Laterality = laterality,
                IncludeLnImn = includeLnImn,
                Medial0AlreadyExists = medial0Exists
            };

            return true;
        }

        private static bool TryParseAngle(string text, out double angle)
        {
            angle = 0.0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            double parsed;
            bool ok = double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed)
                || double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed);

            if (!ok)
                return false;

            if (parsed < 0.0 || parsed > 360.0)
                return false;

            angle = parsed == 360.0 ? 0.0 : parsed;
            return true;
        }

        private static double NormalizeAngle(double angle)
        {
            double normalized = angle % 360.0;
            if (normalized < 0.0)
                normalized += 360.0;
            return normalized;
        }
    }

    internal sealed class BeamAngleDialog : Window
    {
        private readonly TextBox _medial0Text;
        private readonly TextBox _collimatorText;
        private readonly TextBox _couchText;
        private readonly ComboBox _lateralityCombo;
        private readonly CheckBox _includeLnCheck;
        private readonly CheckBox _medial0ExistsCheck;
        private readonly TextBox _previewText;

        public BeamAngleInput SelectedInput { get; private set; }
        public IList<BeamAngleField> PreviewFields { get; private set; }

        public BeamAngleDialog(BeamAngleInput defaults)
        {
            Title = "BreastLRv3 - Flexible Beam Angles";
            Width = 640;
            Height = 520;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanMinimize;

            var root = new Grid { Margin = new Thickness(12) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = root;

            _medial0Text = AddLabeledTextBox(root, 0, "Medial0 Gantry (deg)", defaults.Medial0Gantry.ToString("0.##", CultureInfo.InvariantCulture));
            _collimatorText = AddLabeledTextBox(root, 1, "Collimator (deg, optional)", defaults.Collimator.ToString("0.##", CultureInfo.InvariantCulture));
            _couchText = AddLabeledTextBox(root, 2, "Couch (deg, optional)", defaults.Couch.ToString("0.##", CultureInfo.InvariantCulture));

            var sidePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 6) };
            sidePanel.Children.Add(new TextBlock { Text = "Laterality", Width = 200, VerticalAlignment = VerticalAlignment.Center });
            _lateralityCombo = new ComboBox { Width = 240 };
            _lateralityCombo.Items.Add("Left");
            _lateralityCombo.Items.Add("Right");
            _lateralityCombo.SelectedIndex = defaults.Laterality == BreastLaterality.Left ? 0 : 1;
            sidePanel.Children.Add(_lateralityCombo);
            Grid.SetRow(sidePanel, 3);
            root.Children.Add(sidePanel);

            _includeLnCheck = new CheckBox
            {
                Margin = new Thickness(0, 4, 0, 4),
                Content = "Include LN/IMN fields (12-field template)",
                IsChecked = defaults.IncludeLnImn
            };
            Grid.SetRow(_includeLnCheck, 4);
            root.Children.Add(_includeLnCheck);

            _medial0ExistsCheck = new CheckBox
            {
                Margin = new Thickness(0, 4, 0, 8),
                Content = "Medial0 beam already exists (create remaining fields only)",
                IsChecked = defaults.Medial0AlreadyExists
            };
            Grid.SetRow(_medial0ExistsCheck, 5);
            root.Children.Add(_medial0ExistsCheck);

            _previewText = new TextBox
            {
                AcceptsReturn = true,
                AcceptsTab = false,
                IsReadOnly = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                TextWrapping = TextWrapping.NoWrap,
                MinHeight = 240
            };
            Grid.SetRow(_previewText, 6);
            root.Children.Add(_previewText);

            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };

            var previewButton = new Button { Content = "Preview Angles", Width = 120, Margin = new Thickness(0, 0, 8, 0) };
            previewButton.Click += OnPreviewClicked;
            buttonPanel.Children.Add(previewButton);

            var createButton = new Button { Content = "Create Beams", Width = 120, Margin = new Thickness(0, 0, 8, 0) };
            createButton.Click += OnCreateClicked;
            buttonPanel.Children.Add(createButton);

            var cancelButton = new Button { Content = "Cancel", Width = 90 };
            cancelButton.Click += delegate { DialogResult = false; Close(); };
            buttonPanel.Children.Add(cancelButton);

            Grid.SetRow(buttonPanel, 7);
            root.Children.Add(buttonPanel);

            RefreshPreview();
        }

        private static TextBox AddLabeledTextBox(Grid root, int row, string label, string value)
        {
            var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 6) };
            panel.Children.Add(new TextBlock { Text = label, Width = 200, VerticalAlignment = VerticalAlignment.Center });
            var textBox = new TextBox { Width = 240, Text = value };
            panel.Children.Add(textBox);
            Grid.SetRow(panel, row);
            root.Children.Add(panel);
            return textBox;
        }

        private void OnPreviewClicked(object sender, RoutedEventArgs e)
        {
            RefreshPreview();
        }

        private void OnCreateClicked(object sender, RoutedEventArgs e)
        {
            if (!RefreshPreview())
                return;

            DialogResult = true;
            Close();
        }

        private bool RefreshPreview()
        {
            BeamAngleInput input;
            string error;

            if (!BreastBeamAngleCalculator.TryValidateInput(
                _medial0Text.Text,
                _collimatorText.Text,
                _couchText.Text,
                _lateralityCombo.SelectedIndex == 0 ? BreastLaterality.Left : BreastLaterality.Right,
                _includeLnCheck.IsChecked.HasValue && _includeLnCheck.IsChecked.Value,
                _medial0ExistsCheck.IsChecked.HasValue && _medial0ExistsCheck.IsChecked.Value,
                out input,
                out error))
            {
                _previewText.Text = "Validation error: " + error;
                return false;
            }

            var fields = BreastBeamAngleCalculator.Calculate(input);
            var sb = new StringBuilder();
            sb.AppendLine(input.IncludeLnImn ? "12-field template (with LN/IMN)" : "9-field template (breast only)");
            sb.AppendLine(input.Laterality == BreastLaterality.Left ? "Laterality: Left" : "Laterality: Right");
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "Reference Medial0: {0:0.##}°", input.Medial0Gantry));
            sb.AppendLine();
            sb.AppendLine("Field\tOffset\tGantry\tCollimator\tCouch");

            for (int i = 0; i < fields.Count; i++)
            {
                var f = fields[i];
                string marker = (input.Medial0AlreadyExists && f.OffsetFromMedial0 == 0.0) ? " (existing)" : string.Empty;
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "{0}\t{1:+0.##;-0.##;0}°\t{2:0.##}°\t{3:0.##}°\t{4:0.##}°{5}",
                    f.Name,
                    f.OffsetFromMedial0,
                    f.Gantry,
                    f.Collimator,
                    f.Couch,
                    marker));
            }

            _previewText.Text = sb.ToString();
            SelectedInput = input;
            PreviewFields = fields;
            return true;
        }
    }
}
