using Eto.Forms;
using Eto.Drawing;
using Rhino;
using Rhino.DocObjects;
using Rhino.Geometry;
using Rhino.Input.Custom;
using Rhino.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Moments_of_Inertia.Views
{
    [System.Runtime.InteropServices.Guid("284ae1ed-c3aa-45c4-9167-7934731c712f")]
    public class InertiaPropertiesPanel : Panel
    {
        private DropDown materialDropdown;
        private TextBox densityTextBox;
        
        // Outline curve controls
        private Button assignOutlineButton;
        private GridView outlineCurveGridView;
        
        // Hollow curve controls
        private Button assignHollowsButton;
        private GridView hollowCurvesGridView;
        
        private DropDown unitDropdown;
        private CheckBox showCentroidCheckBox;
        private CheckBox showAxesCheckBox;
        private Button calculateButton;
        private CheckBox realtimeCheckBox;
        private Button exportButton;
        private Button changeColorsButton;
        private Button showValuesButton;
        private GridView resultsGridView;
        
        // Speichern der ausgewählten Kurven
        private Curve outlineCurve;
        private ObjRef outlineCurveRef;
        private List<Curve> hollowCurves = new List<Curve>();
        private List<ObjRef> hollowCurveRefs = new List<ObjRef>();
        
        // Ergebnisse
        private double area;
        private Point3d centroid;
        private double Ix, Iy;
        private double Wx, Wy;
        private double ix, iy;
        private double mass; // Masse basierend auf Dichte

        // Visualisierungsobjekte
        private Guid centroidPointId = Guid.Empty;
        private Guid xAxisId = Guid.Empty;
        private Guid yAxisId = Guid.Empty;
        private Guid textElementId = Guid.Empty;
        
        // Status-Flag für Farbänderung
        private bool useCustomColors = false;

        // Material-Dichte-Dictionary
        private Dictionary<string, double> materialDensities = new Dictionary<string, double>
        {
            { "Steel", 7.85 },     // g/cm³
            { "Aluminum", 2.7 },   // g/cm³
            { "Wood", 0.7 },       // g/cm³
            { "Concrete", 2.4 },   // g/cm³
            { "Glass", 2.5 },      // g/cm³
            { "Custom", 1.0 }      // g/cm³
        };

        // Timer für regelmäßige Aktualisierungen
        private UITimer _updateTimer;

        public InertiaPropertiesPanel()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            // Material selection
            materialDropdown = new DropDown();
            foreach (var material in materialDensities.Keys)
            {
                materialDropdown.Items.Add(material);
            }
            materialDropdown.SelectedIndex = 0;
            materialDropdown.SelectedIndexChanged += MaterialDropdown_SelectedIndexChanged;

            // Density input
            densityTextBox = new TextBox
            {
                Text = materialDensities["Steel"].ToString(),
                ReadOnly = true
            };

            // Outline curve controls with improved layout
            assignOutlineButton = new Button
            {
                Text = "Assign Outline",
                Width = 120
            };
            assignOutlineButton.Click += AssignOutlineButton_Click;
            
            // GridView für Außenkurve-Eigenschaften
            outlineCurveGridView = new GridView
            {
                Height = 100
            };
            SetupCurvePropertiesGrid(outlineCurveGridView);

            // Hollow curves controls with improved layout
            assignHollowsButton = new Button
            {
                Text = "Assign Hollows",
                Width = 120
            };
            assignHollowsButton.Click += AssignHollowsButton_Click;
            
            // GridView für Hohlraumkurven-Eigenschaften
            hollowCurvesGridView = new GridView
            {
                Height = 100
            };
            SetupCurvePropertiesGrid(hollowCurvesGridView);

            // Units dropdown
            unitDropdown = new DropDown
            {
                Items = { "mm", "cm" }
            };
            unitDropdown.SelectedIndex = 1; // Default to cm statt mm

            // Visualization options
            showCentroidCheckBox = new CheckBox
            {
                Text = "Show Centroid"
            };
            showCentroidCheckBox.CheckedChanged += VisualizationOption_Changed;

            showAxesCheckBox = new CheckBox
            {
                Text = "Show Principal Axes"
            };
            showAxesCheckBox.CheckedChanged += VisualizationOption_Changed;

            // Calculate button
            calculateButton = new Button
            {
                Text = "Calculate"
            };
            calculateButton.Click += CalculateButton_Click;

            // Realtime checkbox statt Button
            realtimeCheckBox = new CheckBox
            {
                Text = "Realtime Update",
                ToolTip = "Toggle realtime update of values when curves change"
            };
            realtimeCheckBox.CheckedChanged += RealtimeCheckBox_CheckedChanged;

            // Export button
            exportButton = new Button
            {
                Text = "Export Results",
                Enabled = false
            };
            exportButton.Click += ExportButton_Click;

            // Change Colors button
            changeColorsButton = new Button
            {
                Text = "Change Colors",
                ToolTip = "Toggle between custom colors and layer colors"
            };
            changeColorsButton.Click += ChangeColorsButton_Click;

            // Show Values as TextElement button
            showValuesButton = new Button
            {
                Text = "Show Values as Text",
                ToolTip = "Display values as text element in Rhino next to the profile",
                Enabled = false
            };
            showValuesButton.Click += ShowValuesButton_Click;

            // Results grid
            resultsGridView = new GridView
            {
                Height = 200
            };
            SetupResultsGrid();

            // Layout
            var layout = new DynamicLayout { Padding = new Padding(10) };
            
            // Material section
            layout.AddRow(new Label { Text = "Material Properties:", Font = new Eto.Drawing.Font(SystemFont.Bold) });
            layout.AddRow(new Label { Text = "Material:" }, materialDropdown);
            layout.AddRow(new Label { Text = "Density (g/cm³):" }, densityTextBox);
            layout.AddRow(null); // Spacer
            
            // Curve selection section
            layout.AddRow(new Label { Text = "Curve Selection:", Font = new Eto.Drawing.Font(SystemFont.Bold) });
            
            // Outline curve layout 
            var outlineLayout = new TableLayout
            {
                Padding = new Padding(0),
                Spacing = new Size(5, 5),
                Rows = { 
                    new TableRow(
                        new TableCell(assignOutlineButton, true)
                    ),
                    new TableRow(
                        new TableCell(outlineCurveGridView, true)
                    )
                }
            };
            layout.AddRow(outlineLayout);
            
            // Hollow curve layout
            var hollowLayout = new TableLayout
            {
                Padding = new Padding(0),
                Spacing = new Size(5, 5),
                Rows = { 
                    new TableRow(
                        new TableCell(assignHollowsButton, true)
                    ),
                    new TableRow(
                        new TableCell(hollowCurvesGridView, true)
                    )
                }
            };
            layout.AddRow(hollowLayout);
            layout.AddRow(null); // Spacer
            
            // Options section
            layout.AddRow(new Label { Text = "Options:", Font = new Eto.Drawing.Font(SystemFont.Bold) });
            layout.AddRow(new Label { Text = "Units:" }, unitDropdown);
            layout.AddRow(showCentroidCheckBox);
            layout.AddRow(showAxesCheckBox);
            layout.AddRow(null); // Spacer
            
            // Action buttons
            var buttonLayout = new TableLayout
            {
                Padding = new Padding(0),
                Spacing = new Size(5, 0),
                Rows = { new TableRow(calculateButton, exportButton, changeColorsButton, showValuesButton) }
            };
            layout.AddRow(buttonLayout);
            layout.AddRow(realtimeCheckBox);
            layout.AddRow(null); // Spacer
            
            // Results section
            layout.AddRow(new Label { Text = "Results:", Font = new Eto.Drawing.Font(SystemFont.Bold) });
            layout.AddRow(resultsGridView);

            Content = layout;
        }

        private void MaterialDropdown_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedMaterial = materialDropdown.SelectedKey;
            if (selectedMaterial == "Custom")
            {
                densityTextBox.ReadOnly = false;
            }
            else
            {
                densityTextBox.ReadOnly = true;
                densityTextBox.Text = materialDensities[selectedMaterial].ToString();
            }
        }

        private void SetupCurvePropertiesGrid(GridView gridView)
        {
            gridView.Columns.Add(new GridColumn
            {
                HeaderText = "Property",
                DataCell = new TextBoxCell { Binding = Binding.Property<CurvePropertyItem, string>(r => r.Property) },
                Width = 120
            });

            gridView.Columns.Add(new GridColumn
            {
                HeaderText = "Value",
                DataCell = new TextBoxCell { Binding = Binding.Property<CurvePropertyItem, string>(r => r.Value) }
            });

            // Initial empty data
            gridView.DataStore = new List<CurvePropertyItem>();
        }

        private void SetupResultsGrid()
        {
            resultsGridView.Columns.Add(new GridColumn
            {
                HeaderText = "Property",
                DataCell = new TextBoxCell { Binding = Binding.Property<ResultItem, string>(r => r.Property) }
            });

            resultsGridView.Columns.Add(new GridColumn
            {
                HeaderText = "Value",
                DataCell = new TextBoxCell { Binding = Binding.Property<ResultItem, string>(r => r.Value) }
            });

            resultsGridView.Columns.Add(new GridColumn
            {
                HeaderText = "Unit",
                DataCell = new TextBoxCell { Binding = Binding.Property<ResultItem, string>(r => r.Unit) }
            });

            // Initial empty data
            resultsGridView.DataStore = new List<ResultItem>();
        }

        private void UpdateCurveInfo()
        {
            // Update outline curve info
            var outlineProperties = new List<CurvePropertyItem>();
            
            if (outlineCurve != null)
            {
                // Name (falls vorhanden)
                if (outlineCurveRef != null)
                {
                    RhinoObject obj = outlineCurveRef.Object();
                    if (obj != null)
                    {
                        string name = obj.Name;
                        if (!string.IsNullOrEmpty(name))
                        {
                            outlineProperties.Add(new CurvePropertyItem("Name", name));
                        }
                        
                        outlineProperties.Add(new CurvePropertyItem("ID", obj.Id.ToString()));
                        
                        try
                        {
                            Layer layer = obj.Document.Layers[obj.Attributes.LayerIndex];
                            outlineProperties.Add(new CurvePropertyItem("Layer", layer.Name));
                        }
                        catch
                        {
                            outlineProperties.Add(new CurvePropertyItem("Layer", obj.Attributes.LayerIndex.ToString()));
                        }
                    }
                }
                
                outlineProperties.Add(new CurvePropertyItem("Type", outlineCurve.GetType().Name));
                
                // Einheiten für die Länge basierend auf ausgewählter Einheit hinzufügen
                string unit = unitDropdown.SelectedKey;
                double factor = unit == "mm" ? 1.0 : 0.1; // mm zu cm Umrechnung
                
                outlineProperties.Add(new CurvePropertyItem("Length", $"{outlineCurve.GetLength() * factor:F2} {unit}"));
                outlineProperties.Add(new CurvePropertyItem("Closed", outlineCurve.IsClosed.ToString()));
                
                BoundingBox bbox = outlineCurve.GetBoundingBox(true);
                double width = bbox.Max.X - bbox.Min.X;
                double height = bbox.Max.Y - bbox.Min.Y;
                outlineProperties.Add(new CurvePropertyItem("Size", $"{width * factor:F2} x {height * factor:F2} {unit}"));                
            }
            
            outlineCurveGridView.DataStore = outlineProperties;
            
            // Update hollow curves info
            var hollowProperties = new List<CurvePropertyItem>();
            
            if (hollowCurves.Count > 0)
            {
                hollowProperties.Add(new CurvePropertyItem("Count", hollowCurves.Count.ToString()));
                
                // Einheiten für die Länge basierend auf ausgewählter Einheit hinzufügen
                string unit = unitDropdown.SelectedKey;
                double factor = unit == "mm" ? 1.0 : 0.1; // mm zu cm Umrechnung
                
                double totalLength = 0;
                for (int i = 0; i < hollowCurves.Count; i++)
                {
                    totalLength += hollowCurves[i].GetLength();
                }
                
                hollowProperties.Add(new CurvePropertyItem("Total Length", $"{totalLength * factor:F2} {unit}"));
                
                // Einzelne Kurveninformationen
                for (int i = 0; i < Math.Min(hollowCurves.Count, 10); i++)
                {
                    if (i < hollowCurveRefs.Count && hollowCurveRefs[i] != null)
                    {
                        RhinoObject obj = hollowCurveRefs[i].Object();
                        if (obj != null)
                        {
                            string name = obj.Name;
                            if (!string.IsNullOrEmpty(name))
                            {
                                hollowProperties.Add(new CurvePropertyItem($"Name {i+1}", name));
                            }
                            
                            hollowProperties.Add(new CurvePropertyItem($"ID {i+1}", obj.Id.ToString()));
                        }
                    }
                    
                    hollowProperties.Add(new CurvePropertyItem($"Length {i+1}", $"{hollowCurves[i].GetLength() * factor:F2} {unit}"));
                }
                
                if (hollowCurves.Count > 10)
                {
                    hollowProperties.Add(new CurvePropertyItem("Note", $"{hollowCurves.Count - 10} more curve(s) not shown"));
                }
            }
            
            hollowCurvesGridView.DataStore = hollowProperties;
        }

        private void VisualizationOption_Changed(object sender, EventArgs e)
        {
            if (centroid != Point3d.Unset)
            {
                UpdateVisualization();
            }
        }

        private void UpdateVisualization()
        {
            RhinoDoc doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;

            // Remove existing visualization objects
            if (centroidPointId != Guid.Empty)
            {
                doc.Objects.Delete(centroidPointId, true);
                centroidPointId = Guid.Empty;
            }

            if (xAxisId != Guid.Empty)
            {
                doc.Objects.Delete(xAxisId, true);
                xAxisId = Guid.Empty;
            }

            if (yAxisId != Guid.Empty)
            {
                doc.Objects.Delete(yAxisId, true);
                yAxisId = Guid.Empty;
            }
            
            if (textElementId != Guid.Empty)
            {
                doc.Objects.Delete(textElementId, true);
                textElementId = Guid.Empty;
            }

            // Add centroid point if checked
            if (showCentroidCheckBox.Checked.GetValueOrDefault() && centroid != Point3d.Unset)
            {
                centroidPointId = doc.Objects.AddPoint(centroid);
                
                // Set point display properties
                RhinoObject pointObj = doc.Objects.Find(centroidPointId);
                if (pointObj != null)
                {
                    pointObj.Attributes.ColorSource = ObjectColorSource.ColorFromObject;
                    pointObj.Attributes.ObjectColor = System.Drawing.Color.Red;
                    pointObj.CommitChanges();
                }
                
                doc.Objects.Select(centroidPointId);
            }

            // Add principal axes if checked
            if (showAxesCheckBox.Checked.GetValueOrDefault() && centroid != Point3d.Unset)
            {
                // Determine an appropriate axis length based on the boundary size
                double axisLength = 10;
                if (outlineCurve != null)
                {
                    BoundingBox bbox = outlineCurve.GetBoundingBox(true);
                    double width = bbox.Max.X - bbox.Min.X;
                    double height = bbox.Max.Y - bbox.Min.Y;
                    axisLength = Math.Max(width, height) * 0.25; // 25% of the larger dimension
                }
            
                // X-Axis (red)
                Line xAxis = new Line(centroid, new Point3d(centroid.X + axisLength, centroid.Y, centroid.Z));
                xAxisId = doc.Objects.AddLine(xAxis);
                
                // Set line display properties
                RhinoObject xAxisObj = doc.Objects.Find(xAxisId);
                if (xAxisObj != null)
                {
                    xAxisObj.Attributes.ColorSource = ObjectColorSource.ColorFromObject;
                    xAxisObj.Attributes.ObjectColor = System.Drawing.Color.Red;
                    xAxisObj.CommitChanges();
                }
                
                doc.Objects.Select(xAxisId);

                // Y-Axis (green)
                Line yAxis = new Line(centroid, new Point3d(centroid.X, centroid.Y + axisLength, centroid.Z));
                yAxisId = doc.Objects.AddLine(yAxis);
                
                // Set line display properties
                RhinoObject yAxisObj = doc.Objects.Find(yAxisId);
                if (yAxisObj != null)
                {
                    yAxisObj.Attributes.ColorSource = ObjectColorSource.ColorFromObject;
                    yAxisObj.Attributes.ObjectColor = System.Drawing.Color.Green;
                    yAxisObj.CommitChanges();
                }
                
                doc.Objects.Select(yAxisId);
            }

            doc.Views.Redraw();
        }

        private void AssignOutlineButton_Click(object sender, EventArgs e)
        {
            RhinoDoc doc = RhinoDoc.ActiveDoc;
            if (doc == null)
            {
                MessageBox.Show("No active document", "Error");
                return;
            }

            // Rhino-Befehl ausführen, um eine Kurve auszuwählen
            GetObject go = new GetObject();
            go.SetCommandPrompt("Select outline curve");
            go.GeometryFilter = ObjectType.Curve;
            go.GeometryAttributeFilter = GeometryAttributeFilter.ClosedCurve;
            go.Get();

            if (go.CommandResult() != Rhino.Commands.Result.Success)
            {
                MessageBox.Show("Outline curve selection canceled", "Information");
                return;
            }

            ObjRef objRef = go.Object(0);
            Curve curve = objRef.Curve();
            
            if (curve == null)
            {
                MessageBox.Show("Invalid curve selected", "Error");
                return;
            }

            if (!curve.IsClosed)
            {
                MessageBox.Show("Selected curve must be closed", "Error");
                return;
            }

            if (!curve.IsPlanar())
            {
                MessageBox.Show("Selected curve must be planar", "Error");
                return;
            }

            // Kurve und Referenz speichern
            outlineCurve = curve.DuplicateCurve();
            outlineCurveRef = objRef;
            
            // Update curve info
            UpdateCurveInfo();
            
            // Clear previous results
            ClearResults();
            
            MessageBox.Show("Outline curve assigned", "Success");
        }

        private void AssignHollowsButton_Click(object sender, EventArgs e)
        {
            RhinoDoc doc = RhinoDoc.ActiveDoc;
            if (doc == null)
            {
                MessageBox.Show("No active document", "Error");
                return;
            }

            if (outlineCurve == null)
            {
                MessageBox.Show("Please assign outline curve first", "Error");
                return;
            }

            // Rhino-Befehl ausführen, um mehrere Kurven auszuwählen
            GetObject go = new GetObject();
            go.SetCommandPrompt("Select hollow curves");
            go.GeometryFilter = ObjectType.Curve;
            go.GeometryAttributeFilter = GeometryAttributeFilter.ClosedCurve;
            go.EnablePreSelect(false, true);
            go.GetMultiple(1, 0);

            if (go.CommandResult() != Rhino.Commands.Result.Success)
            {
                MessageBox.Show("Hollow curves selection canceled", "Information");
                return;
            }

            // Bestehende Hohlräume löschen
            hollowCurves.Clear();
            hollowCurveRefs.Clear();

            // Neue Hohlräume hinzufügen
            for (int i = 0; i < go.ObjectCount; i++)
            {
                ObjRef objRef = go.Object(i);
                Curve curve = objRef.Curve();
                
                if (curve == null || !curve.IsClosed || !curve.IsPlanar())
                    continue;

                // Prüfen, ob die Hohlraumkurve innerhalb der Umrisskurve liegt
                if (!IsInsideOutline(curve))
                {
                    MessageBox.Show("All hollow curves must be inside the outline", "Error");
                    hollowCurves.Clear();
                    hollowCurveRefs.Clear();
                    return;
                }

                hollowCurves.Add(curve.DuplicateCurve());
                hollowCurveRefs.Add(objRef);
            }
            
            // Update curve info
            UpdateCurveInfo();
            
            // Clear previous results
            ClearResults();
            
            MessageBox.Show($"{hollowCurves.Count} hollow curve(s) assigned", "Success");
        }

        private void ClearResults()
        {
            resultsGridView.DataStore = new List<ResultItem>();
            exportButton.Enabled = false;
            
            // Auch den ShowValues-Button deaktivieren
            showValuesButton.Enabled = false;
            
            // Clear visualization
            centroid = Point3d.Unset;
            UpdateVisualization();
        }

        private bool IsInsideOutline(Curve curve)
        {
            if (outlineCurve == null || curve == null)
                return false;

            // Prüfen, ob die Kurve innerhalb der Umrisskurve liegt
            Plane plane;
            if (!outlineCurve.TryGetPlane(out plane))
                return false;

            // Mittelpunkt der Kurve berechnen
            AreaMassProperties amp = AreaMassProperties.Compute(curve);
            if (amp == null)
                return false;

            Point3d centroid = amp.Centroid;
            
            // Prüfen, ob der Mittelpunkt innerhalb der Umrisskurve liegt
            return outlineCurve.Contains(centroid, plane, 0.001) == PointContainment.Inside;
        }

        private void CalculateButton_Click(object sender, EventArgs e)
        {
            if (outlineCurve == null)
            {
                MessageBox.Show("Please assign outline curve first", "Error");
                return;
            }

            try
            {
                // Erstellen einer planaren Fläche aus der Umrisskurve
                Curve[] boundaries = new Curve[hollowCurves.Count + 1];
                boundaries[0] = outlineCurve;
                for (int i = 0; i < hollowCurves.Count; i++)
                {
                    boundaries[i + 1] = hollowCurves[i];
                }

                // Erstellen einer planaren Fläche mit Hohlräumen
                Brep[] breps = Brep.CreatePlanarBreps(boundaries, 0.001);
                if (breps == null || breps.Length == 0)
                {
                    MessageBox.Show("Failed to create planar surface", "Error");
                    return;
                }

                // Berechnung der Flächeneigenschaften
                AreaMassProperties amp = AreaMassProperties.Compute(breps[0]);
                if (amp == null)
                {
                    MessageBox.Show("Failed to compute area properties", "Error");
                    return;
                }

                // Ergebnisse speichern
                area = amp.Area;
                centroid = amp.Centroid;

                // Berechnung der Trägheitsmomente
                // Wir verwenden die Brep-Geometrie für eine genauere Berechnung
                Ix = 0;
                Iy = 0;

                // Wir teilen das Brep in Flächen auf
                Mesh[] meshes = Mesh.CreateFromBrep(breps[0], MeshingParameters.Default);
                if (meshes == null || meshes.Length == 0)
                {
                    MessageBox.Show("Failed to create mesh for calculations", "Error");
                    return;
                }

                // Für jedes Mesh
                foreach (Mesh mesh in meshes)
                {
                    // Für jedes Dreieck im Mesh
                    for (int i = 0; i < mesh.Faces.Count; i++)
                    {
                        MeshFace face = mesh.Faces[i];
                        Point3d p1 = mesh.Vertices[face.A];
                        Point3d p2 = mesh.Vertices[face.B];
                        Point3d p3 = mesh.Vertices[face.C];

                        // Fläche des Dreiecks
                        double triangleArea = AreaOfTriangle(p1, p2, p3);

                        // Schwerpunkt des Dreiecks
                        Point3d triangleCentroid = new Point3d(
                            (p1.X + p2.X + p3.X) / 3.0,
                            (p1.Y + p2.Y + p3.Y) / 3.0,
                            (p1.Z + p2.Z + p3.Z) / 3.0);

                        // Beitrag zum Trägheitsmoment (bezogen auf den Gesamtschwerpunkt)
                        double dx = triangleCentroid.X - centroid.X;
                        double dy = triangleCentroid.Y - centroid.Y;

                        // Steiner'scher Satz: I = I_cm + m*d²
                        Ix += triangleArea * dy * dy;
                        Iy += triangleArea * dx * dx;
                    }
                }

                // Widerstandsmomente berechnen
                // Vereinfachte Berechnung: Wir nehmen an, dass die Kurve symmetrisch ist
                BoundingBox bbox = outlineCurve.GetBoundingBox(true);
                
                double yMax = Math.Max(Math.Abs(centroid.Y - bbox.Min.Y), Math.Abs(centroid.Y - bbox.Max.Y));
                double xMax = Math.Max(Math.Abs(centroid.X - bbox.Min.X), Math.Abs(centroid.X - bbox.Max.X));
                
                Wx = Ix / yMax;
                Wy = Iy / xMax;
                
                // Trägheitsradien berechnen
                ix = Math.Sqrt(Ix / area);
                iy = Math.Sqrt(Iy / area);
                
                // Masse berechnen (basierend auf Dichte)
                double density = 0;
                if (double.TryParse(densityTextBox.Text, out density))
                {
                    // Da das Dokument standardmäßig in mm ist und wir eine Profillänge von 1 Meter annehmen:
                    // Dichte: g/cm³
                    // Fläche: mm²
                    // Profillänge: 1000 mm
                    
                    // Umrechnung: g/cm³ -> g/mm³ (dividieren durch 1000)
                    double densityInGPerMm3 = density / 1000.0;
                    
                    // Masse = Volumen * Dichte
                    // Volumen = Fläche * Länge (1000 mm)
                    // Umrechnung in kg: dividieren durch 1000
                    mass = area * 1000.0 * densityInGPerMm3 / 1000.0;
                }
                
                // Ergebnisse anzeigen
                DisplayResults();
                
                // Visualisierung aktualisieren
                UpdateVisualization();
                
                // Export-Button aktivieren
                exportButton.Enabled = true;
                
                // ShowValues-Button aktivieren
                showValuesButton.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Calculation Error");
            }
        }

        private void RealtimeCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            RhinoDoc doc = RhinoDoc.ActiveDoc;
            if (doc == null)
            {
                MessageBox.Show("No active document", "Error");
                realtimeCheckBox.Checked = false;
                return;
            }
            
            if (realtimeCheckBox.Checked == true)
            {
                // Prüfen, ob eine Umrisskurve ausgewählt wurde
                if (outlineCurve == null)
                {
                    MessageBox.Show("Please assign outline curve first", "Error");
                    realtimeCheckBox.Checked = false;
                    return;
                }
                
                // Event-Handler für Änderungen an Objekten hinzufügen
                RhinoDoc.AddRhinoObject += Doc_AddRhinoObject;
                RhinoDoc.DeleteRhinoObject += Doc_DeleteRhinoObject;
                
                // Timer für regelmäßige Überprüfung starten
                _updateTimer = new UITimer();
                _updateTimer.Interval = 1.0; // 1 Sekunde
                _updateTimer.Elapsed += Timer_Elapsed;
                _updateTimer.Start();
                
                // Initial berechnen
                CalculateButton_Click(this, EventArgs.Empty);
            }
            else
            {
                // Event-Handler entfernen
                RhinoDoc.AddRhinoObject -= Doc_AddRhinoObject;
                RhinoDoc.DeleteRhinoObject -= Doc_DeleteRhinoObject;
                
                // Timer stoppen und bereinigen
                if (_updateTimer != null)
                {
                    try
                    {
                        _updateTimer.Stop();
                        _updateTimer.Elapsed -= Timer_Elapsed;
                        _updateTimer.Dispose();
                    }
                    catch (Exception) { /* Ignorieren, falls Timer-Objekt bereits freigegeben wurde */ }
                    _updateTimer = null;
                }
            }
        }
        
        // Separater Event-Handler für den Timer, der Exceptions fängt
        private void Timer_Elapsed(object sender, EventArgs e)
        {
            try
            {
                // Prüfen, ob das Dokument noch gültig ist
                RhinoDoc doc = RhinoDoc.ActiveDoc;
                if (doc == null)
                {
                    // Wenn das Dokument nicht mehr existiert, Timer deaktivieren
                    if (_updateTimer != null)
                    {
                        _updateTimer.Stop();
                        return;
                    }
                }
                
                // Nur prüfen und aktualisieren, wenn nötig und alles gültig ist
                if (HasCurveChanged())
                {
                    UpdateCurveInfo();
                    
                    // Nur berechnen, wenn noch eine Outline existiert
                    if (outlineCurve != null)
                    {
                        CalculateButton_Click(this, EventArgs.Empty);
                    }
                    else
                    {
                        // Wenn keine Outline mehr vorhanden ist, Echtzeit-Update deaktivieren
                        realtimeCheckBox.Checked = false;
                    }
                }
            }
            catch (Exception ex)
            {
                // Fehler protokollieren oder anzeigen, je nach Anforderung
                System.Diagnostics.Debug.WriteLine($"Timer error: {ex.Message}");
                
                // Bei schwerwiegenden Fehlern Timer deaktivieren
                try
                {
                    if (_updateTimer != null)
                    {
                        _updateTimer.Stop();
                        realtimeCheckBox.Checked = false;
                    }
                }
                catch { /* Ignorieren */ }
            }
        }
        
        // Event-Handler für hinzugefügte Objekte
        private void Doc_AddRhinoObject(object sender, RhinoObjectEventArgs e)
        {
            // Prüfen, ob es sich um eine Kurve handelt
            if (e.TheObject.Geometry is Curve)
            {
                // Da wir sowieso einen regelmäßigen Timer haben, 
                // lassen wir die nächste Timerprüfung die Aktualisierung durchführen
                // ohne einen neuen Timer zu erstellen
            }
        }
        
        // Event-Handler für gelöschte Objekte
        private void Doc_DeleteRhinoObject(object sender, RhinoObjectEventArgs e)
        {
            // Wir brauchen keine sofortige Reaktion mehr,
            // da der Timer regelmäßig HasCurveChanged aufruft
            // und dort automatisch gelöschte Kurven erkannt werden
        }
        
        // Prüft, ob sich eine der Kurven geändert hat und aktualisiert sie bei Bedarf
        private bool HasCurveChanged()
        {
            try
            {
                RhinoDoc doc = RhinoDoc.ActiveDoc;
                if (doc == null) return false;
                
                // Update-Check basierend auf IDs statt IsDeformable
                bool needsUpdate = false;
                
                // Prüfen, ob die Outline-Kurve noch vorhanden ist
                if (outlineCurveRef != null)
                {
                    try
                    {
                        RhinoObject obj = doc.Objects.Find(outlineCurveRef.ObjectId);
                        if (obj == null)
                        {
                            // Objekt wurde gelöscht
                            outlineCurve = null;
                            outlineCurveRef = null;
                            needsUpdate = true;
                        }
                        else if (obj.Geometry is Curve curve)
                        {
                            // Prüfen, ob sich die Kurve geändert hat
                            if (!GeometriesAreEqual(outlineCurve, curve))
                            {
                                outlineCurve = curve.DuplicateCurve();
                                needsUpdate = true;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Bei Fehlern (z.B. wenn Objekt durch Undo gelöscht wurde) Referenz löschen
                        outlineCurveRef = null;
                        needsUpdate = true;
                    }
                }
                
                // Ähnliche Prüfung für Hohlraumkurven
                for (int i = hollowCurveRefs.Count - 1; i >= 0; i--)
                {
                    if (hollowCurveRefs[i] != null)
                    {
                        try
                        {
                            RhinoObject obj = doc.Objects.Find(hollowCurveRefs[i].ObjectId);
                            if (obj == null)
                            {
                                // Objekt wurde gelöscht
                                hollowCurves.RemoveAt(i);
                                hollowCurveRefs.RemoveAt(i);
                                needsUpdate = true;
                            }
                            else if (obj.Geometry is Curve curve)
                            {
                                // Prüfen, ob sich die Kurve geändert hat
                                if (!GeometriesAreEqual(hollowCurves[i], curve))
                                {
                                    hollowCurves[i] = curve.DuplicateCurve();
                                    needsUpdate = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // Bei Fehlern Referenz entfernen
                            if (i < hollowCurves.Count)
                                hollowCurves.RemoveAt(i);
                            hollowCurveRefs.RemoveAt(i);
                            needsUpdate = true;
                        }
                    }
                }
                
                return needsUpdate;
            }
            catch (Exception)
            {
                // Bei schwerwiegenden Fehlern false zurückgeben
                return false;
            }
        }
        
        // Hilfsmethode zum Vergleichen von Kurven
        private bool GeometriesAreEqual(Curve c1, Curve c2)
        {
            if (c1 == null || c2 == null) return false;
            
            // BoundingBox-Vergleich als schneller Test
            BoundingBox bb1 = c1.GetBoundingBox(true);
            BoundingBox bb2 = c2.GetBoundingBox(true);
            
            if (!bb1.IsValid || !bb2.IsValid) return false;
            
            // Wenn sich die BoundingBox erheblich geändert hat, sind die Kurven unterschiedlich
            if (Math.Abs(bb1.Min.X - bb2.Min.X) > 0.001 ||
                Math.Abs(bb1.Min.Y - bb2.Min.Y) > 0.001 ||
                Math.Abs(bb1.Max.X - bb2.Max.X) > 0.001 ||
                Math.Abs(bb1.Max.Y - bb2.Max.Y) > 0.001)
                return false;
                
            // Vergleich der Kurvenparameter
            if (c1.Degree != c2.Degree || c1.SpanCount != c2.SpanCount)
                return false;
                
            // Vergleich der Kurvenform durch Längenvergleich
            if (Math.Abs(c1.GetLength() - c2.GetLength()) > 0.001)
                return false;
                
            // Wenn alle Tests bestanden wurden, betrachten wir die Kurven als gleich
            return true;
        }

        private void DisplayResults()
        {
            string unit = unitDropdown.SelectedKey;
            double factor = unit == "mm" ? 1.0 : 0.1; // mm zu cm Umrechnung
            
            var results = new List<ResultItem>
            {
                new ResultItem("Area", (area * factor * factor).ToString("F2"), unit + "²"),
                new ResultItem("Centroid X", (centroid.X * factor).ToString("F2"), unit),
                new ResultItem("Centroid Y", (centroid.Y * factor).ToString("F2"), unit),
                new ResultItem("Moment of Inertia Ix", (Ix * Math.Pow(factor, 4)).ToString("F2"), unit + "⁴"),
                new ResultItem("Moment of Inertia Iy", (Iy * Math.Pow(factor, 4)).ToString("F2"), unit + "⁴"),
                new ResultItem("Polar Moment of Inertia Ip", ((Ix + Iy) * Math.Pow(factor, 4)).ToString("F2"), unit + "⁴"),
                new ResultItem("Section Modulus Wx", (Wx * Math.Pow(factor, 3)).ToString("F2"), unit + "³"),
                new ResultItem("Section Modulus Wy", (Wy * Math.Pow(factor, 3)).ToString("F2"), unit + "³"),
                new ResultItem("Radius of Gyration ix", (ix * factor).ToString("F2"), unit),
                new ResultItem("Radius of Gyration iy", (iy * factor).ToString("F2"), unit)
            };
            
            // Umrisskurvenlänge hinzufügen
            if (outlineCurve != null)
            {
                results.Add(new ResultItem("Outline Length", (outlineCurve.GetLength() * factor).ToString("F2"), unit));
            }
            
            // Masse hinzufügen, wenn berechnet
            if (mass > 0)
            {
                results.Add(new ResultItem("Mass", mass.ToString("F2"), "kg/m"));
                
                // Massenträgheitsmomente hinzufügen
                double massIx = Ix * mass / area;
                double massIy = Iy * mass / area;
                results.Add(new ResultItem("Mass Moment of Inertia Ix", (massIx * Math.Pow(factor, 2)).ToString("F2"), "kg·" + unit + "²/m"));
                results.Add(new ResultItem("Mass Moment of Inertia Iy", (massIy * Math.Pow(factor, 2)).ToString("F2"), "kg·" + unit + "²/m"));
            }
            
            resultsGridView.DataStore = results;
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            var saveDialog = new Eto.Forms.SaveFileDialog
            {
                Title = "Export Results",
                Filters = { new FileFilter("CSV Files", "*.csv") }
            };
            
            if (saveDialog.ShowDialog(this) == DialogResult.Ok)
            {
                try
                {
                    using (StreamWriter writer = new StreamWriter(saveDialog.FileName))
                    {
                        writer.WriteLine("Property,Value,Unit");
                        
                        var results = resultsGridView.DataStore as List<ResultItem>;
                        foreach (var item in results)
                        {
                            writer.WriteLine($"{item.Property},{item.Value},{item.Unit}");
                        }
                    }
                    
                    MessageBox.Show("Results exported successfully", "Export Complete");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error exporting results: {ex.Message}", "Export Error");
                }
            }
        }

        private double AreaOfTriangle(Point3d p1, Point3d p2, Point3d p3)
        {
            // Berechnung der Fläche eines Dreiecks mit der Heron-Formel
            double a = p1.DistanceTo(p2);
            double b = p2.DistanceTo(p3);
            double c = p3.DistanceTo(p1);
            
            double s = (a + b + c) / 2.0;
            return Math.Sqrt(s * (s - a) * (s - b) * (s - c));
        }

        private void ChangeColorsButton_Click(object sender, EventArgs e)
        {
            RhinoDoc doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            
            // Umschalten zwischen benutzerdefinierten Farben und Layer-Farben
            useCustomColors = !useCustomColors;
            
            if (useCustomColors)
            {
                // Ändern der Farbe der Umrisskurve auf Blau
                if (outlineCurveRef != null)
                {
                    RhinoObject obj = outlineCurveRef.Object();
                    if (obj != null)
                    {
                        // Auf benutzerdefinierte Farbe umstellen
                        obj.Attributes.ColorSource = ObjectColorSource.ColorFromObject;
                        // Blau
                        obj.Attributes.ObjectColor = System.Drawing.Color.Blue;
                        obj.CommitChanges();
                    }
                }
                
                // Ändern der Farbe der Hohlraumkurven auf Orange
                foreach (ObjRef hollowRef in hollowCurveRefs)
                {
                    if (hollowRef != null)
                    {
                        RhinoObject obj = hollowRef.Object();
                        if (obj != null)
                        {
                            // Auf benutzerdefinierte Farbe umstellen
                            obj.Attributes.ColorSource = ObjectColorSource.ColorFromObject;
                            // Orange
                            obj.Attributes.ObjectColor = System.Drawing.Color.Orange;
                            obj.CommitChanges();
                        }
                    }
                }
                
                // Button-Text aktualisieren
                changeColorsButton.Text = "Reset Colors";
                
                // Nachricht anzeigen
                MessageBox.Show("Custom colors applied:\n- Outline: Blue\n- Hollows: Orange", "Colors Changed");
            }
            else
            {
                // Zurücksetzen der Farbe der Umrisskurve auf Layer-Farbe
                if (outlineCurveRef != null)
                {
                    RhinoObject obj = outlineCurveRef.Object();
                    if (obj != null)
                    {
                        // Auf Layer-Farbe umstellen
                        obj.Attributes.ColorSource = ObjectColorSource.ColorFromLayer;
                        obj.CommitChanges();
                    }
                }
                
                // Zurücksetzen der Farbe der Hohlraumkurven auf Layer-Farbe
                foreach (ObjRef hollowRef in hollowCurveRefs)
                {
                    if (hollowRef != null)
                    {
                        RhinoObject obj = hollowRef.Object();
                        if (obj != null)
                        {
                            // Auf Layer-Farbe umstellen
                            obj.Attributes.ColorSource = ObjectColorSource.ColorFromLayer;
                            obj.CommitChanges();
                        }
                    }
                }
                
                // Button-Text aktualisieren
                changeColorsButton.Text = "Change Colors";
                
                // Nachricht anzeigen
                MessageBox.Show("Curves set to use layer colors", "Colors Reset");
            }
            
            // Aktualisieren der Anzeige
            doc.Views.Redraw();
        }

        private void ShowValuesButton_Click(object sender, EventArgs e)
        {
            RhinoDoc doc = RhinoDoc.ActiveDoc;
            if (doc == null) return;
            
            // Löschen eines bestehenden TextElements
            if (textElementId != Guid.Empty)
            {
                doc.Objects.Delete(textElementId, true);
                textElementId = Guid.Empty;
            }
            
            if (outlineCurve == null || centroid == Point3d.Unset)
            {
                MessageBox.Show("Please calculate values first", "Error");
                return;
            }
            
            // Einheiten holen
            string unit = unitDropdown.SelectedKey;
            double factor = unit == "mm" ? 1.0 : 0.1; // mm zu cm Umrechnung
            
            // Position für das TextElement berechnen
            // Wir platzieren es rechts neben dem Profil
            BoundingBox bbox = outlineCurve.GetBoundingBox(true);
            double textX = bbox.Max.X + bbox.Diagonal.Length * 0.1;
            double textY = bbox.Max.Y - bbox.Diagonal.Length * 0.1;
            Point3d textPosition = new Point3d(textX, textY, 0);
            
            // Text mit den berechneten Werten erstellen
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("===== PROFILE PROPERTIES =====");
            sb.AppendLine($"Area: {(area * factor * factor):F2} {unit}²");
            sb.AppendLine($"Centroid: ({centroid.X * factor:F2}, {centroid.Y * factor:F2}) {unit}");
            sb.AppendLine($"Moment of Inertia Ix: {(Ix * Math.Pow(factor, 4)):F2} {unit}⁴");
            sb.AppendLine($"Moment of Inertia Iy: {(Iy * Math.Pow(factor, 4)):F2} {unit}⁴");
            sb.AppendLine($"Section Modulus Wx: {(Wx * Math.Pow(factor, 3)):F2} {unit}³");
            sb.AppendLine($"Section Modulus Wy: {(Wy * Math.Pow(factor, 3)):F2} {unit}³");
            sb.AppendLine($"Radius of Gyration ix: {(ix * factor):F2} {unit}");
            sb.AppendLine($"Radius of Gyration iy: {(iy * factor):F2} {unit}");
            
            if (outlineCurve != null)
            {
                sb.AppendLine($"Outline Length: {(outlineCurve.GetLength() * factor):F2} {unit}");
            }
            
            if (mass > 0)
            {
                sb.AppendLine($"Mass: {mass:F2} kg/m");
            }
            
            // TextElement in Rhino erstellen
            var textEntity = new TextEntity
            {
                Plane = new Plane(textPosition, Vector3d.ZAxis),
                PlainText = sb.ToString(),
                Justification = TextJustification.Left,
                Font = new Rhino.DocObjects.Font("Arial")
            };
            
            // Schriftgröße basierend auf der Profil-Größe skalieren
            double textHeight = bbox.Diagonal.Length * 0.02;
            textEntity.TextHeight = textHeight > 0 ? textHeight : 1.0;
            
            // TextEntity zur Szene hinzufügen
            textElementId = doc.Objects.AddText(textEntity);
            
            // Eigenschaften des TextElements setzen
            RhinoObject textObj = doc.Objects.Find(textElementId);
            if (textObj != null)
            {
                textObj.Attributes.ColorSource = ObjectColorSource.ColorFromObject;
                textObj.Attributes.ObjectColor = System.Drawing.Color.Black;
                textObj.CommitChanges();
            }
            
            // Ansicht aktualisieren
            doc.Views.Redraw();
            
            MessageBox.Show("Values displayed as text element in Rhino", "Success");
        }
    }

    // Hilfsklasse für die Kurven-Eigenschaften
    public class CurvePropertyItem
    {
        public string Property { get; set; }
        public string Value { get; set; }

        public CurvePropertyItem(string property, string value)
        {
            Property = property;
            Value = value;
        }
    }

    // Hilfsklasse für die Ergebnisanzeige
    public class ResultItem
    {
        public string Property { get; set; }
        public string Value { get; set; }
        public string Unit { get; set; }

        public ResultItem(string property, string value, string unit)
        {
            Property = property;
            Value = value;
            Unit = unit;
        }
    }
} 