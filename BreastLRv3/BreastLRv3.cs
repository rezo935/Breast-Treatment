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
        private const int MaxBeamIdLength = 16;
        private const int PreferredBeamIdBaseLength = 12;

        public void Execute(ScriptContext context)
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
                "Inputs: Medial0={0:0.##}°, Collimator={1:0.##}°, Couch={2:0.##}°, Laterality={3}, LN/IMN={4}, RemainingOnly={5}, Machine={6}, Energy={7}, DoseRate={8}",
                request.Medial0Gantry,
                request.Collimator,
                request.Couch,
                request.Laterality == BreastLaterality.Left ? "Left" : "Right",
                request.IncludeLnImn ? "Yes" : "No",
                request.Medial0AlreadyExists ? "Yes" : "No",
                request.MachineId,
                request.EnergyModeDisplayName,
                request.DoseRate));

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
                input.MachineId,
                input.EnergyModeDisplayName,
                input.DoseRate,
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

            root = root.Length > PreferredBeamIdBaseLength ? root.Substring(0, PreferredBeamIdBaseLength) : root;
            string candidate = root;
            int suffix = 1;

            while (plan.Beams.Any(b => string.Equals(b.Id, candidate, StringComparison.OrdinalIgnoreCase)))
            {
                string suffixText = suffix.ToString(CultureInfo.InvariantCulture);
                int maxBase = Math.Max(1, MaxBeamIdLength - suffixText.Length);
                string trimmedRoot = root.Length > maxBase ? root.Substring(0, maxBase) : root;
                candidate = trimmedRoot + suffixText;
                suffix++;
            }

            return candidate;
        }

        private static BeamAngleInput BuildDefaults(ScriptContext context, ExternalPlanSetup plan)
        {
            var machineEnergies = DiscoverMachineEnergyOptions(context, plan);
            var input = new BeamAngleInput();
            input.Collimator = 0.0;
            input.Couch = 0.0;
            input.Medial0Gantry = 300.0;
            input.Laterality = GuessLaterality(context);
            input.IncludeLnImn = false;
            input.Medial0AlreadyExists = false;
            input.AvailableMachineEnergies = machineEnergies;

            var beam = plan.Beams.FirstOrDefault(b => !b.IsSetupField);
            if (beam != null)
            {
                input.Medial0Gantry = NormalizeAngle(beam.GantryAngle);
                input.Collimator = NormalizeAngle(beam.CollimatorAngle);
                input.Couch = NormalizeAngle(beam.PatientSupportAngle);
                input.MachineId = beam.TreatmentUnit != null ? beam.TreatmentUnit.Id : null;
                input.EnergyModeDisplayName = beam.EnergyModeDisplayName;
                input.DoseRate = beam.DoseRate;
            }

            if (string.IsNullOrWhiteSpace(input.MachineId) && machineEnergies.Count > 0)
            {
                input.MachineId = machineEnergies[0].MachineId;
                input.EnergyModeDisplayName = machineEnergies[0].EnergyModeDisplayName;
                input.DoseRate = machineEnergies[0].DoseRate;
            }

            return input;
        }

        private static List<MachineEnergyOption> DiscoverMachineEnergyOptions(ScriptContext context, ExternalPlanSetup plan)
        {
            var discovered = new Dictionary<Tuple<string, string, int>, MachineEnergyOption>();

            Action<Beam> collect = beam =>
            {
                if (beam == null || beam.IsSetupField || beam.TreatmentUnit == null)
                    return;

                string machineId = beam.TreatmentUnit.Id;
                string energy = beam.EnergyModeDisplayName;
                if (string.IsNullOrWhiteSpace(machineId) || string.IsNullOrWhiteSpace(energy))
                    return;

                var key = Tuple.Create(machineId, energy, beam.DoseRate);
                if (discovered.ContainsKey(key))
                    return;

                discovered.Add(key, new MachineEnergyOption
                {
                    MachineId = machineId,
                    EnergyModeDisplayName = energy,
                    DoseRate = beam.DoseRate
                });
            };

            if (context != null && context.Course != null)
            {
                foreach (var coursePlan in context.Course.ExternalPlanSetups)
                {
                    if (coursePlan == null)
                        continue;

                    foreach (var beam in coursePlan.Beams)
                        collect(beam);
                }
            }

            if (plan != null)
            {
                foreach (var beam in plan.Beams)
                    collect(beam);
            }

            return discovered.Values
                .OrderBy(v => v.MachineId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(v => v.EnergyModeDisplayName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(v => v.DoseRate)
                .ToList();
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
        public string MachineId { get; set; }
        public string EnergyModeDisplayName { get; set; }
        public int DoseRate { get; set; }
        public IList<MachineEnergyOption> AvailableMachineEnergies { get; set; }
    }

    internal sealed class MachineEnergyOption
    {
        public string MachineId { get; set; }
        public string EnergyModeDisplayName { get; set; }
        public int DoseRate { get; set; }
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
        // Small tolerance to safely treat parsed values extremely close to 360° as 360°.
        private const double AngleTolerance = 0.000000001;

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
            MachineEnergyOption selectedMachineEnergy,
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

            if (selectedMachineEnergy == null
                || string.IsNullOrWhiteSpace(selectedMachineEnergy.MachineId)
                || string.IsNullOrWhiteSpace(selectedMachineEnergy.EnergyModeDisplayName))
            {
                error = "Machine and energy selection is required.";
                return false;
            }

            input = new BeamAngleInput
            {
                Medial0Gantry = NormalizeAngle(medial0),
                Collimator = NormalizeAngle(collimator),
                Couch = NormalizeAngle(couch),
                Laterality = laterality,
                IncludeLnImn = includeLnImn,
                Medial0AlreadyExists = medial0Exists,
                MachineId = selectedMachineEnergy.MachineId,
                EnergyModeDisplayName = selectedMachineEnergy.EnergyModeDisplayName,
                DoseRate = selectedMachineEnergy.DoseRate
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

            angle = Math.Abs(parsed - 360.0) < AngleTolerance ? 360.0 : parsed;
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
        private readonly ComboBox _machineCombo;
        private readonly ComboBox _energyCombo;
        private readonly CheckBox _includeLnCheck;
        private readonly CheckBox _medial0ExistsCheck;
        private readonly TextBox _previewText;
        private readonly IList<MachineEnergyOption> _machineEnergyOptions;
        private readonly BeamAngleInput _initialInput;

        public BeamAngleInput SelectedInput { get; private set; }
        public IList<BeamAngleField> PreviewFields { get; private set; }

        public BeamAngleDialog(BeamAngleInput defaults)
        {
            Title = "BreastLRv3 - Flexible Beam Angles";
            Width = 640;
            Height = 600;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanMinimize;
            _initialInput = defaults;
            _machineEnergyOptions = defaults.AvailableMachineEnergies ?? new List<MachineEnergyOption>();

            var root = new Grid { Margin = new Thickness(12) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
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

            var machinePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 6) };
            machinePanel.Children.Add(new TextBlock { Text = "Machine", Width = 200, VerticalAlignment = VerticalAlignment.Center });
            _machineCombo = new ComboBox { Width = 240 };
            var machineIds = _machineEnergyOptions
                .Select(o => o.MachineId)
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();
            for (int i = 0; i < machineIds.Count; i++)
                _machineCombo.Items.Add(machineIds[i]);
            _machineCombo.SelectionChanged += OnMachineSelectionChanged;
            machinePanel.Children.Add(_machineCombo);
            Grid.SetRow(machinePanel, 3);
            root.Children.Add(machinePanel);

            var sidePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 6) };
            sidePanel.Children.Add(new TextBlock { Text = "Laterality", Width = 200, VerticalAlignment = VerticalAlignment.Center });
            _lateralityCombo = new ComboBox { Width = 240 };
            _lateralityCombo.Items.Add("Left");
            _lateralityCombo.Items.Add("Right");
            _lateralityCombo.SelectedIndex = defaults.Laterality == BreastLaterality.Left ? 0 : 1;
            sidePanel.Children.Add(_lateralityCombo);
            Grid.SetRow(sidePanel, 4);
            root.Children.Add(sidePanel);

            var energyPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 6, 0, 6) };
            energyPanel.Children.Add(new TextBlock { Text = "Energy", Width = 200, VerticalAlignment = VerticalAlignment.Center });
            _energyCombo = new ComboBox { Width = 240 };
            energyPanel.Children.Add(_energyCombo);
            Grid.SetRow(energyPanel, 5);
            root.Children.Add(energyPanel);

            _includeLnCheck = new CheckBox
            {
                Margin = new Thickness(0, 4, 0, 4),
                Content = "Include LN/IMN fields (12-field template)",
                IsChecked = defaults.IncludeLnImn
            };
            Grid.SetRow(_includeLnCheck, 6);
            root.Children.Add(_includeLnCheck);

            _medial0ExistsCheck = new CheckBox
            {
                Margin = new Thickness(0, 4, 0, 8),
                Content = "Medial0 beam already exists (create remaining fields only)",
                IsChecked = defaults.Medial0AlreadyExists
            };
            Grid.SetRow(_medial0ExistsCheck, 7);
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
            Grid.SetRow(_previewText, 8);
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

            Grid.SetRow(buttonPanel, 9);
            root.Children.Add(buttonPanel);

            if (_machineCombo.Items.Count > 0)
            {
                int machineIndex = 0;
                if (!string.IsNullOrWhiteSpace(defaults.MachineId))
                {
                    for (int i = 0; i < _machineCombo.Items.Count; i++)
                    {
                        if (string.Equals(_machineCombo.Items[i] as string, defaults.MachineId, StringComparison.OrdinalIgnoreCase))
                        {
                            machineIndex = i;
                            break;
                        }
                    }
                }
                _machineCombo.SelectedIndex = machineIndex;
            }

            RefreshEnergyOptions(defaults);

            RefreshPreview();
        }

        private void OnMachineSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RefreshEnergyOptions(SelectedInput ?? _initialInput);
        }

        private void RefreshEnergyOptions(BeamAngleInput currentInput)
        {
            _energyCombo.Items.Clear();
            string selectedMachine = _machineCombo.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(selectedMachine))
                return;

            var matching = _machineEnergyOptions
                .Where(o => string.Equals(o.MachineId, selectedMachine, StringComparison.OrdinalIgnoreCase))
                .OrderBy(o => o.EnergyModeDisplayName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(o => o.DoseRate)
                .ToList();

            for (int i = 0; i < matching.Count; i++)
            {
                var option = matching[i];
                var item = new ComboBoxItem();
                item.Content = string.Format(CultureInfo.InvariantCulture, "{0} (Dose rate {1})", option.EnergyModeDisplayName, option.DoseRate);
                item.Tag = option;
                _energyCombo.Items.Add(item);
            }

            if (_energyCombo.Items.Count == 0)
                return;

            int selectedIndex = 0;
            if (currentInput != null
                && string.Equals(currentInput.MachineId, selectedMachine, StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(currentInput.EnergyModeDisplayName))
            {
                for (int i = 0; i < _energyCombo.Items.Count; i++)
                {
                    var item = _energyCombo.Items[i] as ComboBoxItem;
                    var option = item != null ? item.Tag as MachineEnergyOption : null;
                    if (option != null
                        && string.Equals(option.EnergyModeDisplayName, currentInput.EnergyModeDisplayName, StringComparison.OrdinalIgnoreCase)
                        && option.DoseRate == currentInput.DoseRate)
                    {
                        selectedIndex = i;
                        break;
                    }
                }
            }

            _energyCombo.SelectedIndex = selectedIndex;
        }

        private MachineEnergyOption GetSelectedMachineEnergy()
        {
            var item = _energyCombo.SelectedItem as ComboBoxItem;
            if (item == null)
                return null;
            return item.Tag as MachineEnergyOption;
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
                GetSelectedMachineEnergy(),
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
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                "Machine/Energy: {0} / {1} (Dose rate {2})", input.MachineId, input.EnergyModeDisplayName, input.DoseRate));
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
