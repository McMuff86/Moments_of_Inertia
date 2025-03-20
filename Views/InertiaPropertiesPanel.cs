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
using System.Threading.Tasks;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;

namespace Moments_of_Inertia.Views
{
    [System.Runtime.InteropServices.Guid("284ae1ed-c3aa-45c4-9167-7934731c712f")]
    public class InertiaPropertiesPanel : Panel
    {
        private DropDown materialDropdown;
        private TextBox densityTextBox;
        private TextBox profileDepthTextBox;
        private TextBox momentXTextBox; // Biegemoment um die X-Achse
        private TextBox momentYTextBox; // Biegemoment um die Y-Achse
        private TextBox yieldStrengthTextBox; // Streckgrenze des Materials
        private NumericStepper utilizationFactorStepper; // Sicherheitsfaktor für die Ausnutzungsberechnung
        private Label utilizationResultLabel; // Anzeige der berechneten Ausnutzung
        
        // Neue Eingabefelder für Querkräfte und Torsion
        private TextBox shearForceXTextBox;
        private TextBox shearForceYTextBox;
        private TextBox torsionTextBox;
        
        // Outline curve controls
        private Button assignOutlineButton;
        private GridView outlineCurveGridView;
        
        // Hollow curve controls
        private Button assignHollowsButton;
        private GridView hollowCurvesGridView;
        
        private DropDown unitDropdown;
        private CheckBox showCentroidCheckBox;
        private CheckBox showAxesCheckBox;
        private CheckBox highAccuracyCheckBox; // Neu: Option für hohe Genauigkeit
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
        
        // Neue Variablen für maximale Abstände vom Schwerpunkt
        private double xMax, yMax;

        // Visualisierungsobjekte
        private Guid centroidPointId = Guid.Empty;
        private Guid xAxisId = Guid.Empty;
        private Guid yAxisId = Guid.Empty;
        private Guid textElementId = Guid.Empty;
        private List<Guid> meshVisualizationIds = new List<Guid>(); // Neue Mesh-Visualisierungsobjekte
        
        // Status-Flag für Farbänderung
        private bool useCustomColors = false;

        // Material-Dichte-Dictionary
        private Dictionary<string, double[]> materialProperties = new Dictionary<string, double[]>
        {
            // Werte: [Dichte (g/cm³), Streckgrenze (N/mm²)]
            { "Steel", new double[] { 7.85, 235.0 } },      // Baustahl S235
            { "Aluminum", new double[] { 2.7, 160.0 } },   // Aluminium-Legierung
            { "Wood", new double[] { 0.7, 20.0 } },        // Holz (durchschnittlich)
            { "Concrete", new double[] { 2.4, 30.0 } },    // Beton (Druckfestigkeit)
            { "Glass", new double[] { 2.5, 50.0 } },       // Glas
            { "Custom", new double[] { 1.0, 100.0 } }      // Benutzerdefiniert
        };

        // Timer für regelmäßige Aktualisierungen
        private UITimer _updateTimer;

        // Speichern der berechneten Ausnutzungswerte
        private Dictionary<string, double> utilizationValues;

        // Neue Checkbox für die Anzeige der Spannungsverteilung
        private CheckBox showStressDistributionCheckBox;

        public InertiaPropertiesPanel()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            // Material selection
            materialDropdown = new DropDown();
            foreach (var material in materialProperties.Keys)
            {
                materialDropdown.Items.Add(material);
            }
            materialDropdown.SelectedIndex = 0;
            materialDropdown.SelectedIndexChanged += MaterialDropdown_SelectedIndexChanged;

            // Density input
            densityTextBox = new TextBox
            {
                Text = materialProperties["Steel"][0].ToString(),
                ReadOnly = true
            };

            // Profile depth input
            profileDepthTextBox = new TextBox
            {
                Text = "1000", // Default 1000 mm = 1 m
                ToolTip = "Enter the profile depth for mass calculation (in mm)"
            };

            // Biegemoment Mx
            momentXTextBox = new TextBox
            {
                Text = "0",
                ToolTip = "Enter bending moment around X-axis (kNm)"
            };

            // Biegemoment My
            momentYTextBox = new TextBox
            {
                Text = "0",
                ToolTip = "Enter bending moment around Y-axis (kNm)"
            };

            // Neue Eingabefelder für Querkräfte und Torsion
            shearForceXTextBox = new TextBox
            {
                Text = "0",
                ToolTip = "Enter shear force in X direction (kN)"
            };

            shearForceYTextBox = new TextBox
            {
                Text = "0",
                ToolTip = "Enter shear force in Y direction (kN)"
            };

            torsionTextBox = new TextBox
            {
                Text = "0",
                ToolTip = "Enter torsion moment (kNm)"
            };

            // Streckgrenze
            yieldStrengthTextBox = new TextBox
            {
                Text = materialProperties["Steel"][1].ToString(),
                ToolTip = "Yield strength of the material (N/mm²)"
            };

            // Sicherheitsfaktor für die Ausnutzung
            utilizationFactorStepper = new NumericStepper
            {
                Value = 1.0,
                MinValue = 0.1,
                MaxValue = 2.0,
                Increment = 0.1,
                DecimalPlaces = 1,
                ToolTip = "Safety factor for utilization calculation"
            };

            // Ergebnis-Label für die Ausnutzung
            utilizationResultLabel = new Label
            {
                Text = "-- %",
                Font = new Eto.Drawing.Font(SystemFont.Bold)
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
                Items = { "mm", "cm", "m" } // Meter als Option hinzugefügt
            };
            unitDropdown.SelectedIndex = 1; // Default to cm

            // High accuracy checkbox
            highAccuracyCheckBox = new CheckBox
            {
                Text = "High Accuracy",
                ToolTip = "Use more precise calculations (slower performance)"
            };

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

            // Neue Checkbox für die Anzeige der Spannungsverteilung
            showStressDistributionCheckBox = new CheckBox
            {
                Text = "Show Stress Distribution",
                Enabled = false // Initial deaktiviert, bis Berechnung erfolgt
            };
            showStressDistributionCheckBox.CheckedChanged += VisualizationOption_Changed;

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
            layout.AddRow(new Label { Text = "Profile Depth (mm):" }, profileDepthTextBox);
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
            layout.AddRow(highAccuracyCheckBox);
            
            // Section Utilization section
            layout.AddRow(new Label { Text = "Section Utilization:", Font = new Eto.Drawing.Font(SystemFont.Bold) });
            layout.AddRow(new Label { Text = "Moment Mx (kNm):" }, momentXTextBox);
            layout.AddRow(new Label { Text = "Moment My (kNm):" }, momentYTextBox);
            layout.AddRow(new Label { Text = "Shear Force Qx (kN):" }, shearForceXTextBox);
            layout.AddRow(new Label { Text = "Shear Force Qy (kN):" }, shearForceYTextBox);
            layout.AddRow(new Label { Text = "Torsion T (kNm):" }, torsionTextBox);
            layout.AddRow(new Label { Text = "Yield Strength (N/mm²):" }, yieldStrengthTextBox);
            layout.AddRow(new Label { Text = "Safety Factor:" }, utilizationFactorStepper);
            layout.AddRow(new Label { Text = "Utilization:" }, utilizationResultLabel);
            
            layout.AddRow(showCentroidCheckBox);
            layout.AddRow(showAxesCheckBox);
            layout.AddRow(showStressDistributionCheckBox);
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
                yieldStrengthTextBox.ReadOnly = false;
            }
            else
            {
                densityTextBox.ReadOnly = true;
                densityTextBox.Text = materialProperties[selectedMaterial][0].ToString();
                
                yieldStrengthTextBox.ReadOnly = true;
                yieldStrengthTextBox.Text = materialProperties[selectedMaterial][1].ToString();
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
                double factor = unit == "mm" ? 1.0 : (unit == "cm" ? 0.1 : 0.001); // Korrigiert für m-Einheit
                
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
                double factor = unit == "mm" ? 1.0 : (unit == "cm" ? 0.1 : 0.001); // Korrigiert für m-Einheit
                
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
            
            // Neue Mesh-Visualisierungsobjekte entfernen
            foreach (Guid id in meshVisualizationIds)
            {
                doc.Objects.Delete(id, true);
            }
            meshVisualizationIds.Clear();

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
            
            // Spannungsverteilung anzeigen, wenn gewünscht
            if (showStressDistributionCheckBox.Checked.GetValueOrDefault() && 
                utilizationValues != null && 
                outlineCurve != null)
            {
                try 
                {
                    // Erstelle ein Mesh für die Visualisierung
                    Curve[] boundaries = new Curve[hollowCurves.Count + 1];
                    boundaries[0] = outlineCurve;
                    for (int i = 0; i < hollowCurves.Count; i++)
                    {
                        boundaries[i + 1] = hollowCurves[i];
                    }
                    
                    // Erstellen der planaren Fläche mit Hohlräumen
                    Brep[] breps = Brep.CreatePlanarBreps(boundaries, 0.001);
                    if (breps != null && breps.Length > 0)
                    {
                        // Höhere Mesh-Dichte für bessere Visualisierung
                        MeshingParameters mp = new MeshingParameters();
                        mp.MinimumEdgeLength = Math.Min(outlineCurve.GetBoundingBox(true).Diagonal.Length / 100, 0.5);
                        mp.MaximumEdgeLength = Math.Min(outlineCurve.GetBoundingBox(true).Diagonal.Length / 50, 2.0);
                        Mesh[] meshes = Mesh.CreateFromBrep(breps[0], mp);
                        
                        if (meshes != null && meshes.Length > 0)
                        {
                            foreach (Mesh mesh in meshes)
                            {
                                // Farben für jeden Vertex basierend auf der Spannung berechnen
                                mesh.VertexColors.CreateMonotoneMesh(System.Drawing.Color.Blue);
                                
                                // Maximale Spannung zur Normalisierung ermitteln
                                double maxStress = utilizationValues["Sigma v"];
                                
                                for (int i = 0; i < mesh.Vertices.Count; i++)
                                {
                                    Point3d vertex = mesh.Vertices[i];
                                    
                                    // Abstand vom Schwerpunkt
                                    double dx = vertex.X - centroid.X;
                                    double dy = vertex.Y - centroid.Y;
                                    
                                    // Spannungen an diesem Punkt berechnen
                                    // Verwende die Klassenvariablen xMax und yMax statt lokale Variablen
                                    double sigmaX = Math.Abs(utilizationValues["Sigma X"] * dy / yMax);
                                    double sigmaY = Math.Abs(utilizationValues["Sigma Y"] * dx / xMax);
                                    double tau = utilizationValues.ContainsKey("Tau") ? utilizationValues["Tau"] : 0;
                                    
                                    // Von-Mises-Spannung
                                    double sigmaV = Math.Sqrt(sigmaX * sigmaX + sigmaY * sigmaY - sigmaX * sigmaY + 3 * tau * tau);
                                    
                                    // Farbwert basierend auf normalisiertem Spannungswert
                                    double normalizedStress = Math.Min(sigmaV / maxStress, 1.0);
                                    
                                    // Farbskala von Blau (0) über Grün (0.5) nach Rot (1.0)
                                    System.Drawing.Color color;
                                    if (normalizedStress < 0.5)
                                    {
                                        // Blau zu Grün
                                        int blue = 255 - (int)(normalizedStress * 2 * 255);
                                        int green = (int)(normalizedStress * 2 * 255);
                                        color = System.Drawing.Color.FromArgb(0, green, blue);
                                    }
                                    else
                                    {
                                        // Grün zu Rot
                                        int green = 255 - (int)((normalizedStress - 0.5) * 2 * 255);
                                        int red = (int)((normalizedStress - 0.5) * 2 * 255);
                                        color = System.Drawing.Color.FromArgb(red, green, 0);
                                    }
                                    
                                    mesh.VertexColors[i] = color;
                                }
                                
                                // Mesh zur Szene hinzufügen
                                Guid meshId = doc.Objects.AddMesh(mesh);
                                meshVisualizationIds.Add(meshId);
                                
                                // Objekt-Eigenschaften setzen
                                RhinoObject meshObj = doc.Objects.Find(meshId);
                                if (meshObj != null)
                                {
                                    meshObj.Attributes.ColorSource = ObjectColorSource.ColorFromObject;
                                    meshObj.Attributes.ObjectColor = System.Drawing.Color.Gray; // Basis-Farbe (wird mit Vertex-Farben überschrieben)
                                    meshObj.CommitChanges();
                                }
                            }
                            
                            // Legende für die Farbskala erstellen
                            CreateStressLegend(doc, utilizationValues["Sigma v"]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating stress visualization: {ex.Message}");
                }
            }

            doc.Views.Redraw();
        }
        
        // Neue Methode zur Erstellung einer Legende für die Spannungsverteilung
        private void CreateStressLegend(RhinoDoc doc, double maxStress)
        {
            try
            {
                // Obere rechte Ecke des Bildschirms bestimmen
                BoundingBox bbox = outlineCurve.GetBoundingBox(true);
                Point3d legendStart = new Point3d(bbox.Max.X + bbox.Diagonal.Length * 0.15, bbox.Min.Y, 0);
                
                // Legende-Dimensionen
                double legendHeight = bbox.Diagonal.Length * 0.3;
                double legendWidth = bbox.Diagonal.Length * 0.05;
                
                // Farbbalken erstellen (10 Segmente)
                int segments = 10;
                for (int i = 0; i < segments; i++)
                {
                    double normalizedStress = (double)i / segments;
                    double nextNormalizedStress = (double)(i + 1) / segments;
                    
                    // Eckpunkte des Rechtecks
                    Point3d corner1 = new Point3d(
                        legendStart.X, 
                        legendStart.Y + normalizedStress * legendHeight, 
                        0);
                        
                    Point3d corner2 = new Point3d(
                        legendStart.X + legendWidth, 
                        legendStart.Y + nextNormalizedStress * legendHeight, 
                        0);
                    
                    // Rechteck erstellen
                    Rectangle3d rect = new Rectangle3d(
                        Plane.WorldXY, 
                        corner1, 
                        corner2);
                    
                    // Farbwert für dieses Segment
                    System.Drawing.Color color;
                    double colorValue = normalizedStress + 0.5 / segments; // Mittelpunkt des Segments
                    
                    if (colorValue < 0.5)
                    {
                        // Blau zu Grün
                        int blue = 255 - (int)(colorValue * 2 * 255);
                        int green = (int)(colorValue * 2 * 255);
                        color = System.Drawing.Color.FromArgb(0, green, blue);
                    }
                    else
                    {
                        // Grün zu Rot
                        int green = 255 - (int)((colorValue - 0.5) * 2 * 255);
                        int red = (int)((colorValue - 0.5) * 2 * 255);
                        color = System.Drawing.Color.FromArgb(red, green, 0);
                    }
                    
                    // Geschlossene Kurve erstellen für Hatch
                    Curve rectCurve = rect.ToNurbsCurve();
                    
                    // Sowohl Umriss als auch gefüllte Fläche hinzufügen
                    
                    // 1. Gefüllte Fläche (Hatch) hinzufügen
                    var hatches = Hatch.Create(rectCurve, 0, 0, 1.0, 0.001);
                    if (hatches != null && hatches.Length > 0)
                    {
                        // Hatch zur Szene hinzufügen
                        Guid hatchId = doc.Objects.AddHatch(hatches[0]);
                        meshVisualizationIds.Add(hatchId);
                        
                        // Farbeigenschaften des Hatch setzen
                        RhinoObject hatchObj = doc.Objects.Find(hatchId);
                        if (hatchObj != null)
                        {
                            hatchObj.Attributes.ColorSource = ObjectColorSource.ColorFromObject;
                            hatchObj.Attributes.ObjectColor = color;
                            hatchObj.CommitChanges();
                        }
                    }
                    
                    // 2. Umriss hinzufügen mit schwarzer Farbe für Kontrast
                    Guid curveId = doc.Objects.AddCurve(rectCurve);
                    meshVisualizationIds.Add(curveId);
                    
                    // Objekteigenschaften für den Umriss setzen
                    RhinoObject obj = doc.Objects.Find(curveId);
                    if (obj != null)
                    {
                        obj.Attributes.ColorSource = ObjectColorSource.ColorFromObject;
                        obj.Attributes.ObjectColor = System.Drawing.Color.Black; // Schwarzer Umriss für besseren Kontrast
                        obj.Attributes.PlotWeight = 0.15; // Dünnere Linie für den Umriss
                        obj.CommitChanges();
                    }
                    
                    // Beschriftungen für Min und Max hinzufügen
                    if (i == 0 || i == segments - 1)
                    {
                        string text = i == 0 ? "0" : $"{maxStress:F0} N/mm²";
                        Point3d textPoint = new Point3d(
                            legendStart.X + legendWidth * 1.2, 
                            legendStart.Y + normalizedStress * legendHeight + (i == segments - 1 ? 0 : legendHeight / segments), 
                            0);
                            
                        var textEntity = new TextEntity
                        {
                            Plane = new Plane(textPoint, Vector3d.ZAxis),
                            PlainText = text,
                            Justification = TextJustification.Left,
                            Font = new Rhino.DocObjects.Font("Arial")
                        };
                        
                        double textHeight = bbox.Diagonal.Length * 0.015;
                        textEntity.TextHeight = textHeight > 0 ? textHeight : 1.0;
                        
                        Guid textId = doc.Objects.AddText(textEntity);
                        meshVisualizationIds.Add(textId);
                    }
                }
                
                // Titel für die Legende
                Point3d titlePoint = new Point3d(
                    legendStart.X, 
                    legendStart.Y - bbox.Diagonal.Length * 0.03, 
                    0);
                    
                var titleEntity = new TextEntity
                {
                    Plane = new Plane(titlePoint, Vector3d.ZAxis),
                    PlainText = "Stress (N/mm²)",
                    Justification = TextJustification.Left,
                    // Einfacher Font ohne Bold-Versuch
                    Font = new Rhino.DocObjects.Font("Arial")
                };
                
                // Größere Schrift für den Titel verwenden
                double titleHeight = bbox.Diagonal.Length * 0.025; // Etwas größer als normale Beschriftung
                titleEntity.TextHeight = titleHeight > 0 ? titleHeight : 1.0;
                
                Guid titleId = doc.Objects.AddText(titleEntity);
                meshVisualizationIds.Add(titleId);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating legend: {ex.Message}");
            }
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

            // Temporäre Liste zur Prüfung vor dem endgültigen Hinzufügen
            List<Curve> tempHollowCurves = new List<Curve>();
            List<ObjRef> tempHollowRefs = new List<ObjRef>();

            // Neue Hohlräume hinzufügen
            for (int i = 0; i < go.ObjectCount; i++)
            {
                ObjRef objRef = go.Object(i);
                Curve curve = objRef.Curve();
                
                if (curve == null || !curve.IsClosed || !curve.IsPlanar())
                {
                    MessageBox.Show($"Curve {i+1} is not a valid closed planar curve and will be ignored.", "Warning");
                    continue;
                }

                // Prüfen, ob die Hohlraumkurve innerhalb der Umrisskurve liegt
                if (!IsInsideOutline(curve))
                {
                    MessageBox.Show($"Curve {i+1} is not completely inside the outline and will be ignored.", "Warning");
                    continue;
                }

                // Prüfen, ob sich die neue Kurve mit bereits ausgewählten Hohlraumkurven überschneidet
                bool hasIntersection = false;
                foreach (Curve existingCurve in tempHollowCurves)
                {
                    if (DoCurvesIntersect(curve, existingCurve))
                    {
                        MessageBox.Show($"Curve {i+1} intersects with another selected hollow curve and will be ignored.", "Warning");
                        hasIntersection = true;
                        break;
                    }
                }

                if (!hasIntersection)
                {
                    tempHollowCurves.Add(curve.DuplicateCurve());
                    tempHollowRefs.Add(objRef);
                }
            }

            // Alle gültigen Kurven zur endgültigen Liste hinzufügen
            hollowCurves = tempHollowCurves;
            hollowCurveRefs = tempHollowRefs;
            
            // Update curve info
            UpdateCurveInfo();
            
            // Clear previous results
            ClearResults();
            
            if (hollowCurves.Count > 0)
            {
                MessageBox.Show($"{hollowCurves.Count} valid hollow curve(s) assigned", "Success");
            }
            else
            {
                MessageBox.Show("No valid hollow curves were assigned. Please ensure curves are closed, planar, inside the outline, and do not intersect with each other.", "Information");
            }
        }

        // Überprüfen ob zwei Kurven sich überschneiden oder eine in der anderen liegt
        private bool DoCurvesIntersect(Curve curve1, Curve curve2)
        {
            // Wenn die Kurven sich überschneiden
            Rhino.Geometry.Intersect.CurveIntersections intersections = 
                Rhino.Geometry.Intersect.Intersection.CurveCurve(curve1, curve2, 0.001, 0.001);

            if (intersections != null && intersections.Count > 0)
                return true;

            // Überprüfen, ob der Schwerpunkt einer Kurve innerhalb der anderen liegt
            AreaMassProperties amp1 = AreaMassProperties.Compute(curve1);
            AreaMassProperties amp2 = AreaMassProperties.Compute(curve2);

            if (amp1 != null && amp2 != null)
            {
                Point3d centroid1 = amp1.Centroid;
                Point3d centroid2 = amp2.Centroid;

                // Korrekte Ebenen für jede Kurve bestimmen
                Plane plane1, plane2;
                if (!curve1.TryGetPlane(out plane1))
                    plane1 = Plane.WorldXY; // Fallback, falls keine Ebene bestimmt werden kann
                
                if (!curve2.TryGetPlane(out plane2))
                    plane2 = Plane.WorldXY; // Fallback, falls keine Ebene bestimmt werden kann

                // Jede Kurve in ihrer eigenen Ebene prüfen
                if (curve1.Contains(centroid2, plane1, 0.001) != PointContainment.Outside ||
                    curve2.Contains(centroid1, plane2, 0.001) != PointContainment.Outside)
                    return true;
            }

            return false;
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

            try
            {
                // Prüfen, ob die Kurve planar ist und die korrekte Ebene ermitteln
                Plane plane;
                if (!outlineCurve.TryGetPlane(out plane))
                    return false;

                // VERBESSERT: Einfache BoundingBox-Prüfung als schneller Ausschlusstest
                BoundingBox curveBBox = curve.GetBoundingBox(true);
                BoundingBox outlineBBox = outlineCurve.GetBoundingBox(true);
                
                // Mit etwas Toleranz für numerische Ungenauigkeiten
                outlineBBox.Inflate(0.001); // Kleine Toleranz hinzufügen

                // Schneller Ausschlusstest: Wenn die Hohlraum-BoundingBox nicht innerhalb
                // der erweiterten Outline-BoundingBox liegt, kann die Kurve nicht vollständig innerhalb liegen
                if (!outlineBBox.Contains(curveBBox))
                    return false;
                
                // Strategiewechsel: Prüfen ob Punkte auf/innerhalb Outline liegen
                
                // 1. Prüfen des Schwerpunkts (für einfache Fälle)
                AreaMassProperties amp = AreaMassProperties.Compute(curve);
                if (amp == null)
                    return false;
                    
                Point3d centroid = amp.Centroid;
                PointContainment centroidContainment = outlineCurve.Contains(centroid, plane, 0.001);
                if (centroidContainment == PointContainment.Outside)
                    return false;
                
                // 2. Unterschiedliche Strategien basierend auf Genauigkeitseinstellung
                if (highAccuracyCheckBox.Checked.GetValueOrDefault())
                {
                    // A. Punktweise Prüfung mit hoher Punktdichte
                    int pointCount = Math.Max(50, (int)(curve.GetLength() / 1.0)); // Viele Punkte für hohe Genauigkeit
                    bool allPointsInside = true;
                    
                    for (int i = 0; i < pointCount; i++)
                    {
                        double t = i / (double)(pointCount - 1);
                        Point3d point = curve.PointAt(t);
                        PointContainment containment = outlineCurve.Contains(point, plane, 0.001); // Korrekte Plane
                        
                        if (containment == PointContainment.Outside)
                        {
                            allPointsInside = false;
                            break;
                        }
                    }
                    
                    if (!allPointsInside)
                        return false;
                    
                    // B. Versuch mit Boolean-Operationen (nur als zusätzlicher Test)
                    try
                    {
                        // Wenn die Kurve vollständig innerhalb liegt, sollte eine
                        // Boolean-Differenz möglich sein und zu einer geschlossenen Kurve führen
                        Curve[] diff = Curve.CreateBooleanDifference(outlineCurve, curve, 0.001);
                        
                        // Wenn keine Differenz erstellt werden konnte, ist etwas falsch
                        if (diff == null || diff.Length == 0)
                            return false;
                        
                        // Es muss mindestens eine geschlossene Kurve als Ergebnis geben
                        bool hasClosedCurve = false;
                        foreach (Curve c in diff)
                        {
                            if (c.IsClosed)
                            {
                                hasClosedCurve = true;
                                break;
                            }
                        }
                        
                        if (!hasClosedCurve)
                            return false;
                        
                        // Berechnung der Flächendifferenz (sollte positiv sein)
                        AreaMassProperties outlineAmp = AreaMassProperties.Compute(outlineCurve);
                        AreaMassProperties curveAmp = AreaMassProperties.Compute(curve);
                        
                        if (outlineAmp != null && curveAmp != null)
                        {
                            // Die Differenz der Flächen sollte relevant sein
                            double areaDiff = outlineAmp.Area - curveAmp.Area;
                            if (areaDiff < -0.001) // Toleranz hinzugefügt
                                return false;
                        }
                    }
                    catch (Exception)
                    {
                        // Bei Fehlern in der Boolean-Operation ignorieren und
                        // Entscheidung auf Basis der Punkttests treffen
                    }
                    
                    // Im High-Accuracy-Modus sind die bisherigen Tests ausreichend
                    // Region-Test wird nicht benötigt
                    return true;
                }
                else
                {
                    // Einfachere Prüfung für Performance-Modus: Einige Punkte auf der Kurve testen
                    int pointCount = Math.Max(8, (int)(curve.GetLength() / 10.0)); // Weniger Punkte im Performance-Modus
                    int outsidePoints = 0;
                    
                    for (int i = 0; i < pointCount; i++)
                    {
                        double t = i / (double)(pointCount - 1);
                        Point3d point = curve.PointAt(t);
                        if (outlineCurve.Contains(point, plane, 0.001) == PointContainment.Outside) // Korrekte Plane
                        {
                            outsidePoints++;
                        }
                    }
                    
                    // Tolerieren von maximal einem Punkt, der als außerhalb erkannt wird (für numerische Stabilität)
                    if (outsidePoints > 1)
                        return false;
                        
                    // Im Performance-Modus wird der Region-Test als zusätzliche Absicherung beibehalten
                    try
                    {
                        // Region aus der Hohlkurve erstellen
                        Curve[] loops = new Curve[] { curve };
                        // Die Methode CreatePlanarBreps statt CreatePlanarFace verwenden
                        Brep[] faces = Brep.CreatePlanarBreps(loops, 0.001);
                        
                        if (faces != null && faces.Length > 0)
                        {
                            // Das erste Brep prüfen
                            foreach (BrepFace face in faces[0].Faces)
                            {
                                // To3dCurve() korrekt verwenden - in ein Array konvertieren falls nötig
                                var loopCurve = face.OuterLoop.To3dCurve();
                                
                                // Immer als einzelne Curve behandeln und in ein Array konvertieren
                                Curve[] regionCurves = new Curve[] { loopCurve };
                                
                                if (regionCurves != null && regionCurves.Length > 0)
                                {
                                    // Ist die Kurve der Region innerhalb der Outline?
                                    bool isValid = true;
                                    foreach (Curve regionCurve in regionCurves)
                                    {
                                        int testPointCount = Math.Max(4, (int)(regionCurve.GetLength() / 20.0));
                                        for (int i = 0; i < testPointCount; i++)
                                        {
                                            double t = i / (double)(testPointCount - 1);
                                            Point3d point = regionCurve.PointAt(t);
                                            
                                            if (outlineCurve.Contains(point, plane, 0.001) == PointContainment.Outside) // Korrekte Plane
                                            {
                                                isValid = false;
                                                break;
                                            }
                                        }
                                        
                                        if (!isValid)
                                            break;
                                    }
                                    
                                    // Wenn die Region ungültig ist, ist die Kurve nicht innen
                                    if (!isValid)
                                        return false;
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Fehler ignorieren - andere Tests könnten trotzdem bestanden werden
                    }
                }
                
                // Wenn alle Tests erfolgreich waren (oder keine Fehler geworfen haben),
                // nehmen wir an, dass die Kurve vollständig innerhalb liegt
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void CalculateButton_Click(object sender, EventArgs e)
        {
            if (outlineCurve == null)
            {
                MessageBox.Show("Please assign outline curve first", "Error");
                return;
            }

            // Ladeindikator anzeigen
            string originalButtonText = calculateButton.Text;
            calculateButton.Text = "Calculating...";
            calculateButton.Enabled = false;
            Application.Instance.RunIteration(); // Aktualisierung der UI erzwingen

            try
            {
                // Validierung der Dichte, falls "Custom" ausgewählt ist
                if (materialDropdown.SelectedKey == "Custom")
                {
                    if (!double.TryParse(densityTextBox.Text, out double customDensity) || customDensity <= 0)
                    {
                        MessageBox.Show("Please enter a positive density value", "Validation Error");
                        calculateButton.Text = originalButtonText;
                        calculateButton.Enabled = true;
                        return;
                    }
                }

                // Validierung der Profiltiefe
                double profileDepth = 1000; // Default: 1000 mm = 1 m
                if (!double.TryParse(profileDepthTextBox.Text, out profileDepth) || profileDepth <= 0)
                {
                    MessageBox.Show("Please enter a positive profile depth value in mm", "Validation Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }
                
                // Obergrenze für sinnvolle Profiltiefe (10 m = 10000 mm)
                if (profileDepth > 10000)
                {
                    MessageBox.Show("Profile depth exceeds 10 m (10000 mm). Values may be unrealistic.", "Warning");
                }
                
                // Validierung der Biegemomente
                double momentX = 0;
                double momentY = 0;
                double yieldStrength = 0;
                double safetyFactor = 1.0;
                // Neue Variablen für Querkräfte und Torsion
                double Qx = 0;
                double Qy = 0;
                double T = 0;
                
                if (!double.TryParse(momentXTextBox.Text, out momentX))
                {
                    MessageBox.Show("Invalid value for moment Mx. Please enter a valid number.", "Validation Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }
                
                if (!double.TryParse(momentYTextBox.Text, out momentY))
                {
                    MessageBox.Show("Invalid value for moment My. Please enter a valid number.", "Validation Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }
                
                // Validierung der Querkräfte und Torsion
                if (!double.TryParse(shearForceXTextBox.Text, out Qx))
                {
                    MessageBox.Show("Invalid value for shear force Qx. Please enter a valid number.", "Validation Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }
                
                if (!double.TryParse(shearForceYTextBox.Text, out Qy))
                {
                    MessageBox.Show("Invalid value for shear force Qy. Please enter a valid number.", "Validation Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }
                
                if (!double.TryParse(torsionTextBox.Text, out T))
                {
                    MessageBox.Show("Invalid value for torsion T. Please enter a valid number.", "Validation Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }
                
                if (!double.TryParse(yieldStrengthTextBox.Text, out yieldStrength) || yieldStrength <= 0)
                {
                    MessageBox.Show("Please enter a positive yield strength value", "Validation Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }
                
                safetyFactor = utilizationFactorStepper.Value;
                if (safetyFactor <= 0)
                {
                    MessageBox.Show("Safety factor must be positive", "Validation Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }

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
                    MessageBox.Show("Failed to create planar surface. Please check if hollow curves are fully inside outline and do not intersect each other.", "Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }

                // Berechnung der Flächeneigenschaften
                AreaMassProperties amp = AreaMassProperties.Compute(breps[0]);
                if (amp == null)
                {
                    MessageBox.Show("Failed to compute area properties", "Error");
                    calculateButton.Text = originalButtonText;
                    calculateButton.Enabled = true;
                    return;
                }

                // Ergebnisse speichern
                area = amp.Area;
                centroid = amp.Centroid;

                // Trägheitsmomente berechnen mit der direkten Rhino-Methode
                // Nutze CentroidCoordinatesMomentsOfInertia für Momente bezogen auf den Schwerpunkt
                Vector3d momentsOfInertia = amp.CentroidCoordinatesMomentsOfInertia;
                
                // Die Momente von Inertia in CentroidCoordinatesMomentsOfInertia sind:
                // X = Ix (Moment um die X-Achse durch den Schwerpunkt)
                // Y = Iy (Moment um die Y-Achse durch den Schwerpunkt)
                // Z = Iz (Moment um die Z-Achse durch den Schwerpunkt)
                Ix = momentsOfInertia.X;
                Iy = momentsOfInertia.Y;
                
                // OPTIMIERT: Zweistufige Berechnung von xMax und yMax für bessere Performance
                // Verwende die Klassenvariablen statt lokaler Variablen
                this.xMax = 0.0;
                this.yMax = 0.0;
                
                // Stufe 1: Schnelle Abschätzung mit BoundingBox
                BoundingBox bbox = outlineCurve.GetBoundingBox(true);
                this.yMax = Math.Max(Math.Abs(bbox.Max.Y - centroid.Y), Math.Abs(bbox.Min.Y - centroid.Y));
                this.xMax = Math.Max(Math.Abs(bbox.Max.X - centroid.X), Math.Abs(bbox.Min.X - centroid.X));
                
                // Stufe 2: Präzise Berechnung mit Punkten auf der Kurve, wenn hohe Genauigkeit gewünscht ist
                if (highAccuracyCheckBox.Checked.GetValueOrDefault())
                {
                    // Prüfen der Randpunkte der Outline-Kurve
                    int outlinePointCount = Math.Max(200, (int)(outlineCurve.GetLength() / 0.5)); // Hohe Dichte
                    for (int i = 0; i < outlinePointCount; i++)
                    {
                        double t = i / (double)(outlinePointCount - 1);
                        Point3d point = outlineCurve.PointAt(t);
                        
                        double dx = Math.Abs(point.X - centroid.X);
                        double dy = Math.Abs(point.Y - centroid.Y);
                        
                        this.xMax = Math.Max(this.xMax, dx);
                        this.yMax = Math.Max(this.yMax, dy);
                    }
                    
                    // Auch Hohlraumkurven prüfen, falls vorhanden
                    foreach (Curve hollowCurve in hollowCurves)
                    {
                        int hollowPointCount = Math.Max(100, (int)(hollowCurve.GetLength() / 0.5)); // Hohe Dichte
                        for (int i = 0; i < hollowPointCount; i++)
                        {
                            double t = i / (double)(hollowPointCount - 1);
                            Point3d point = hollowCurve.PointAt(t);
                            
                            double dx = Math.Abs(point.X - centroid.X);
                            double dy = Math.Abs(point.Y - centroid.Y);
                            
                            this.xMax = Math.Max(this.xMax, dx);
                            this.yMax = Math.Max(this.yMax, dy);
                        }
                    }
                }
                else
                {
                    // Nur bei komplexeren Profilen mit Hohlräumen zusätzlich einfache Stichproben nehmen
                    if (hollowCurves.Count > 0)
                    {
                        // Outline mit geringer Punktdichte prüfen
                        int outlinePointCount = Math.Max(50, (int)(outlineCurve.GetLength() / 2.0));
                        for (int i = 0; i < outlinePointCount; i++)
                        {
                            double t = i / (double)(outlinePointCount - 1);
                            Point3d point = outlineCurve.PointAt(t);
                            
                            double dx = Math.Abs(point.X - centroid.X);
                            double dy = Math.Abs(point.Y - centroid.Y);
                            
                            this.xMax = Math.Max(this.xMax, dx);
                            this.yMax = Math.Max(this.yMax, dy);
                        }
                        
                        // Hohlräume mit geringer Punktdichte prüfen
                        foreach (Curve hollowCurve in hollowCurves)
                        {
                            int hollowPointCount = Math.Max(20, (int)(hollowCurve.GetLength() / 3.0));
                            for (int i = 0; i < hollowPointCount; i++)
                            {
                                double t = i / (double)(hollowPointCount - 1);
                                Point3d point = hollowCurve.PointAt(t);
                                
                                double dx = Math.Abs(point.X - centroid.X);
                                double dy = Math.Abs(point.Y - centroid.Y);
                                
                                this.xMax = Math.Max(this.xMax, dx);
                                this.yMax = Math.Max(this.yMax, dy);
                            }
                        }
                    }
                }
                
                // VERBESSERT: Sicherheitscheck, dass die Abstände nicht Null sind
                if (this.yMax < 0.001)
                {
                    MessageBox.Show("Maximum y-distance from centroid too small for accurate calculations", "Warning");
                    this.yMax = 0.001;
                }
                
                if (this.xMax < 0.001)
                {
                    MessageBox.Show("Maximum x-distance from centroid too small for accurate calculations", "Warning");
                    this.xMax = 0.001;
                }
                
                // Widerstandsmomente berechnen mit genaueren Maximalwerten
                Wx = Ix / this.yMax;
                Wy = Iy / this.xMax;
                
                // Trägheitsradien berechnen
                ix = Math.Sqrt(Ix / area);
                iy = Math.Sqrt(Iy / area);
                
                // Masse berechnen (basierend auf Dichte)
                double materialDensity = 0;
                if (double.TryParse(densityTextBox.Text, out materialDensity))
                {
                    // VERBESSERT: Klare Dokumentation der Annahmen:
                    // - Dichte in g/cm³
                    // - Fläche in mm²
                    // - Profiltiefe in mm (benutzerdefiniert)
                    // - Berechnung: Masse [kg] = Fläche [mm²] * Tiefe [mm] * Dichte [g/cm³] / 10⁶
                    
                    // Umrechnung: g/cm³ -> g/mm³ (dividieren durch 1000)
                    double densityInGPerMm3 = materialDensity / 1000.0;
                    
                    // Masse = Volumen * Dichte
                    // Volumen = Fläche * Tiefe
                    // Umrechnung in kg: dividieren durch 1000
                    mass = area * profileDepth * densityInGPerMm3 / 1000.0;
                }
                else
                {
                    // Wenn die Dichte nicht gültig ist, keine Masse berechnen
                    mass = 0;
                }
                
                // NEU: Mesh-Diskretisierung für spätere FEA-Erweiterung
                List<MeshElementResult> elementResults = new List<MeshElementResult>();
                if (highAccuracyCheckBox.Checked.GetValueOrDefault())
                {
                    // Nur im High-Accuracy-Modus ein feineres Mesh erstellen
                    MeshingParameters mp = new MeshingParameters();
                    mp.MinimumEdgeLength = Math.Min(bbox.Diagonal.Length / 50, 1.0); // Feineres Mesh
                    mp.MaximumEdgeLength = Math.Min(bbox.Diagonal.Length / 20, 5.0);
                    Mesh[] meshes = Mesh.CreateFromBrep(breps[0], mp);
                    
                    if (meshes != null && meshes.Length > 0)
                    {
                        // Durch die Mesh-Vertices iterieren und Spannungen berechnen
                        foreach (Mesh mesh in meshes)
                        {
                            for (int i = 0; i < mesh.Vertices.Count; i++)
                            {
                                Point3d vertex = mesh.Vertices[i];
                                
                                // Abstand vom Schwerpunkt
                                double dx = vertex.X - centroid.X;
                                double dy = vertex.Y - centroid.Y;
                                
                                // Element-Ergebnis für zukünftige FEA-Erweiterung speichern
                                MeshElementResult elementResult = new MeshElementResult
                                {
                                    Centroid = vertex
                                    // Spannungen werden später berechnet
                                };
                                elementResults.Add(elementResult);
                            }
                        }
                    }
                }
                
                // Zusätzlich: Berechnung der Ausnutzung, wenn Biegemomente eingegeben wurden
                if ((momentX != 0) || (momentY != 0) || (Qx != 0) || (Qy != 0) || (T != 0))
                {
                    if (yieldStrength > 0)
                    {
                        // Momente von kNm in Nmm umrechnen (x 10^6)
                        momentX *= 1000000;
                        momentY *= 1000000;
                        
                        // Querkräfte von kN in N umrechnen (x 1000)
                        Qx *= 1000;
                        Qy *= 1000;
                        
                        // Torsionsmoment von kNm in Nmm umrechnen (x 10^6)
                        T *= 1000000;
                        
                        // Normalspannungen berechnen (N/mm²)
                        // Berücksichtigung der Richtung: Positive und negative Spannungen berücksichtigen
                        double sigmaXPos = momentX > 0 ? momentX / Wx : 0;
                        double sigmaXNeg = momentX < 0 ? -momentX / Wx : 0;
                        double sigmaYPos = momentY > 0 ? momentY / Wy : 0;
                        double sigmaYNeg = momentY < 0 ? -momentY / Wy : 0;
                        
                        // Maximale Spannung in jeder Richtung auswählen
                        double sigmaX = Math.Max(Math.Abs(sigmaXPos), Math.Abs(sigmaXNeg));
                        double sigmaY = Math.Max(Math.Abs(sigmaYPos), Math.Abs(sigmaYNeg));
                        
                        // NEU: Berechnung der Schubspannungen
                        // Vereinfachte Schubspannungsberechnung (Näherung)
                        double Jt = Ix + Iy; // Torsionswiderstand (vereinfacht)
                        
                        // Schubspannungen aus Querkräften (approximiert)
                        double width = bbox.Max.X - bbox.Min.X;
                        double height = bbox.Max.Y - bbox.Min.Y;
                        
                        double tauQx = Math.Abs(Qx) > 0.001 ? 1.5 * Math.Abs(Qx) / area : 0;
                        double tauQy = Math.Abs(Qy) > 0.001 ? 1.5 * Math.Abs(Qy) / area : 0;
                        
                        // Schubspannungen aus Torsion (vereinfacht)
                        double tauT = 0;
                        if (Math.Abs(T) > 0.001)
                        {
                            // Vereinfachter Ansatz für geschlossene Querschnitte
                            double maxDistance = Math.Max(this.xMax, this.yMax);
                            tauT = T * maxDistance / Jt;
                        }
                        
                        // Resultierender Schubspannungsvektor
                        double tau = Math.Sqrt(tauQx * tauQx + tauQy * tauQy + tauT * tauT);
                        
                        // Von Mises Vergleichsspannung mit Math.NET berechnen
                        // Spannungstensor (vereinfacht 2D)
                        var stressVector = DenseVector.OfArray(new double[] { sigmaX, sigmaY, 0 });
                        var shearVector = DenseVector.OfArray(new double[] { tau, 0, 0 });
                        
                        // Von Mises: sqrt(sigma_x^2 + sigma_y^2 - sigma_x*sigma_y + 3*tau^2)
                        double sigmaV = Math.Sqrt(
                            stressVector[0] * stressVector[0] + 
                            stressVector[1] * stressVector[1] - 
                            stressVector[0] * stressVector[1] + 
                            3 * shearVector[0] * shearVector[0]);
                        
                        // Ausnutzung berechnen (unter Berücksichtigung des Sicherheitsfaktors)
                        double utilization = sigmaV / (yieldStrength / safetyFactor) * 100;
                        
                        // Ergebnis anzeigen
                        utilizationResultLabel.Text = $"{utilization:F1} %";
                        utilizationResultLabel.TextColor = utilization > 100 ? Colors.Red : Colors.Black;
                        
                        // Auch in die Ergebnistabelle aufnehmen
                        utilizationValues = new Dictionary<string, double>
                        {
                            { "Sigma X", sigmaX },
                            { "Sigma Y", sigmaY },
                            { "Tau Qx", tauQx },
                            { "Tau Qy", tauQy },
                            { "Tau T", tauT },
                            { "Tau", tau },
                            { "Sigma v", sigmaV },
                            { "Utilization", utilization }
                        };
                        
                        // Wenn im High-Accuracy-Modus, Spannungen auch für Mesh-Elemente berechnen
                        if (highAccuracyCheckBox.Checked.GetValueOrDefault() && elementResults.Count > 0)
                        {
                            foreach (MeshElementResult element in elementResults)
                            {
                                // Abstand vom Schwerpunkt
                                double dx = element.Centroid.X - centroid.X;
                                double dy = element.Centroid.Y - centroid.Y;
                                
                                // Spannungen an diesem Punkt
                                element.SigmaX = momentY * dy / Ix;
                                element.SigmaY = momentX * dx / Iy;
                                element.TauXY = Math.Sqrt(tauQx * tauQx + tauQy * tauQy);
                                element.TauT = tauT;
                                element.Tau = tau;
                                
                                // Von-Mises-Spannung
                                element.SigmaV = Math.Sqrt(
                                    element.SigmaX * element.SigmaX + 
                                    element.SigmaY * element.SigmaY - 
                                    element.SigmaX * element.SigmaY + 
                                    3 * element.Tau * element.Tau);
                            }
                        }
                        
                        // Zusatz-Info zur Schubspannung
                        if (Qx != 0 || Qy != 0 || T != 0)
                        {
                            MessageBox.Show(
                                "Note: The calculation now includes approximate values for shear stresses.\n" +
                                "- Shear stresses from Qx and Qy use simplified methods.\n" +
                                "- Torsional shear stresses are approximated for closed sections.\n" +
                                "- For more accurate results, a full FEA analysis is recommended.", 
                                "Stress Calculation Info", 
                                MessageBoxType.Information);
                        }
                    }
                    else
                    {
                        utilizationResultLabel.Text = "Invalid yield strength";
                        utilizationResultLabel.TextColor = Colors.Red;
                        utilizationValues = null;
                    }
                }
                else
                {
                    utilizationResultLabel.Text = "-- %";
                    utilizationResultLabel.TextColor = Colors.Black;
                    utilizationValues = null;
                }
                
                // Ergebnisse anzeigen
                DisplayResults();
                
                // Visualisierung aktualisieren
                UpdateVisualization();
                
                // Export-Button aktivieren
                exportButton.Enabled = true;
                
                // ShowValues-Button aktivieren
                showValuesButton.Enabled = true;
                
                // NEU: Aktiviere die Checkbox für die Spannungsvisualisierung
                showStressDistributionCheckBox.Enabled = utilizationValues != null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Calculation Error");
            }
            finally
            {
                // Ladeindikator entfernen und Button wiederherstellen
                calculateButton.Text = originalButtonText;
                calculateButton.Enabled = true;
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
                
                // Warnung bei aktiviertem High-Accuracy-Modus
                if (highAccuracyCheckBox.Checked.GetValueOrDefault())
                {
                    if (MessageBox.Show(
                        "High accuracy mode is enabled, which may cause performance issues with real-time updates.\n\n" +
                        "Do you want to continue with real-time updates in high accuracy mode?",
                        "Performance Warning",
                        MessageBoxButtons.YesNo) == DialogResult.No)
                    {
                        realtimeCheckBox.Checked = false;
                        return;
                    }
                }
                
                // Event-Handler für Änderungen an Objekten hinzufügen
                RhinoDoc.AddRhinoObject += Doc_AddRhinoObject;
                RhinoDoc.DeleteRhinoObject += Doc_DeleteRhinoObject;
                
                // Timer für regelmäßige Überprüfung starten
                _updateTimer = new UITimer();
                _updateTimer.Interval = highAccuracyCheckBox.Checked.GetValueOrDefault() ? 2.0 : 1.0; // Längeres Intervall bei hoher Genauigkeit
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
        
        // Optimierter Timer_Elapsed mit asynchroner Berechnung
        private async void Timer_Elapsed(object sender, EventArgs e)
        {
            try
            {
                RhinoDoc doc = RhinoDoc.ActiveDoc;
                if (doc == null)
                {
                    if (_updateTimer != null)
                    {
                        _updateTimer.Stop();
                        return;
                    }
                }
                
                bool hasChanged = HasCurveChanged();
                
                if (hasChanged)
                {
                    if (outlineCurve != null)
                    {
                        // Alle Berechnungen und UI-Updates in einem asynchronen Task
                        await Task.Run(() => {
                            try
                            {
                                // Alle UI-Updates in einem einzigen Invoke-Block zusammenfassen
                                Eto.Forms.Application.Instance.Invoke(() => {
                                    try 
                                    {
                                        // Button-Status setzen
                                        calculateButton.Text = "Calculating...";
                                        calculateButton.Enabled = false;
                                        
                                        // Kurveninformationen aktualisieren
                                        UpdateCurveInfo();
                                        
                                        // Berechnung durchführen
                                        CalculateButton_Click(this, EventArgs.Empty);
                                        
                                        // Button-Status zurücksetzen
                                        calculateButton.Text = "Calculate";
                                        calculateButton.Enabled = true;
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"UI update error: {ex.Message}");
                                        calculateButton.Text = "Calculate";
                                        calculateButton.Enabled = true;
                                    }
                                });
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Async calculation error: {ex.Message}");
                                
                                // Fehlerfall: Button Zustand zurücksetzen
                                Eto.Forms.Application.Instance.Invoke(() => {
                                    calculateButton.Text = "Calculate";
                                    calculateButton.Enabled = true;
                                });
                            }
                        });
                    }
                    else
                    {
                        Eto.Forms.Application.Instance.Invoke(() => {
                            realtimeCheckBox.Checked = false;
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Timer error: {ex.Message}");
                
                try
                {
                    if (_updateTimer != null)
                    {
                        _updateTimer.Stop();
                        Eto.Forms.Application.Instance.Invoke(() => {
                            realtimeCheckBox.Checked = false;
                            calculateButton.Text = "Calculate";
                            calculateButton.Enabled = true;
                        });
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
            double factor = 1.0;
            
            // Umrechnungsfaktor für die gewählte Einheit
            if (unit == "cm") 
                factor = 0.1;  // mm zu cm
            else if (unit == "m")
                factor = 0.001; // mm zu m
            
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
                double profileDepth = 1000;
                double.TryParse(profileDepthTextBox.Text, out profileDepth);
                
                // Angepasst für benutzerdefinierte Tiefe
                string massUnit = "kg";
                string depthInfo = "";
                
                if (profileDepth == 1000)
                {
                    massUnit = "kg/m";
                    depthInfo = "for 1m length";
                }
                else
                {
                    depthInfo = $"for {profileDepth}mm depth";
                }
                
                results.Add(new ResultItem($"Mass ({depthInfo})", mass.ToString("F2"), massUnit));
                
                // Massenträgheitsmomente hinzufügen
                double massIx = Ix * mass / area;
                double massIy = Iy * mass / area;
                results.Add(new ResultItem("Mass Moment of Inertia Ix", (massIx * Math.Pow(factor, 2)).ToString("F2"), "kg·" + unit + "²"));
                results.Add(new ResultItem("Mass Moment of Inertia Iy", (massIy * Math.Pow(factor, 2)).ToString("F2"), "kg·" + unit + "²"));
            }
            
            // Ausnutzungswerte hinzufügen, wenn vorhanden
            if (utilizationValues != null)
            {
                results.Add(new ResultItem("", "", ""));  // Leerzeile als Trenner
                results.Add(new ResultItem("Section Utilization", "", ""));
                results.Add(new ResultItem("Stress Sigma X", utilizationValues["Sigma X"].ToString("F2"), "N/mm²"));
                results.Add(new ResultItem("Stress Sigma Y", utilizationValues["Sigma Y"].ToString("F2"), "N/mm²"));
                
                // Neue Werte für Schubspannungen hinzufügen
                if (utilizationValues.ContainsKey("Tau Qx"))
                    results.Add(new ResultItem("Shear Stress Qx", utilizationValues["Tau Qx"].ToString("F2"), "N/mm²"));
                
                if (utilizationValues.ContainsKey("Tau Qy"))
                    results.Add(new ResultItem("Shear Stress Qy", utilizationValues["Tau Qy"].ToString("F2"), "N/mm²"));
                
                if (utilizationValues.ContainsKey("Tau T"))
                    results.Add(new ResultItem("Torsional Stress", utilizationValues["Tau T"].ToString("F2"), "N/mm²"));
                
                if (utilizationValues.ContainsKey("Tau"))
                    results.Add(new ResultItem("Combined Shear Stress", utilizationValues["Tau"].ToString("F2"), "N/mm²"));
                
                results.Add(new ResultItem("Equivalent Stress Sigma v", utilizationValues["Sigma v"].ToString("F2"), "N/mm²"));
                
                string utilizationText = utilizationValues["Utilization"].ToString("F1");
                if (utilizationValues["Utilization"] > 100)
                    utilizationText += " (!)";
                    
                results.Add(new ResultItem("Utilization", utilizationText, "%"));
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
            // Berechnung der Fläche mit Kreuzprodukt
            Vector3d v1 = new Vector3d(p2 - p1);
            Vector3d v2 = new Vector3d(p3 - p1);
            Vector3d cross = Vector3d.CrossProduct(v1, v2);
            
            // Hälfte der Länge des Kreuzprodukts ist die Fläche
            return 0.5 * cross.Length;
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
            double factor = unit == "mm" ? 1.0 : (unit == "cm" ? 0.1 : 0.001); // Korrigiert für m-Einheit
            
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
                double profileDepth = 1000;
                double.TryParse(profileDepthTextBox.Text, out profileDepth);
                
                // Angepasst für benutzerdefinierte Tiefe
                string massUnit = "kg";
                string depthInfo = "";
                
                if (profileDepth == 1000)
                {
                    massUnit = "kg/m";
                    depthInfo = "for 1m profile length";
                }
                else
                {
                    depthInfo = $"for {profileDepth}mm profile depth";
                }
                
                sb.AppendLine($"Mass: {mass:F2} {massUnit} ({depthInfo})");
            }
            
            // Auch Ausnutzungswerte im Textblock anzeigen, wenn vorhanden
            if (utilizationValues != null)
            {
                sb.AppendLine("");
                sb.AppendLine("=== SECTION UTILIZATION ===");
                sb.AppendLine($"Stress Sigma X: {utilizationValues["Sigma X"]:F2} N/mm²");
                sb.AppendLine($"Stress Sigma Y: {utilizationValues["Sigma Y"]:F2} N/mm²");
                
                // Neue Werte für Schubspannungen hinzufügen
                if (utilizationValues.ContainsKey("Tau Qx"))
                    sb.AppendLine($"Shear Stress Qx: {utilizationValues["Tau Qx"]:F2} N/mm²");
                
                if (utilizationValues.ContainsKey("Tau Qy"))
                    sb.AppendLine($"Shear Stress Qy: {utilizationValues["Tau Qy"]:F2} N/mm²");
                
                if (utilizationValues.ContainsKey("Tau T"))
                    sb.AppendLine($"Torsional Stress: {utilizationValues["Tau T"]:F2} N/mm²");
                
                if (utilizationValues.ContainsKey("Tau"))
                    sb.AppendLine($"Combined Shear Stress: {utilizationValues["Tau"]:F2} N/mm²");
                
                sb.AppendLine($"Equivalent Stress: {utilizationValues["Sigma v"]:F2} N/mm²");
                
                string utilizationText = utilizationValues["Utilization"].ToString("F1");
                if (utilizationValues["Utilization"] > 100)
                    utilizationText += " (EXCEEDED)";
                    
                sb.AppendLine($"Utilization: {utilizationText} %");
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
    
    // Hilfsklasse für Mesh-Element-Ergebnisse (FEA-Vorbereitung)
    public class MeshElementResult
    {
        public Point3d Centroid { get; set; } // Schwerpunkt des Elements
        public double SigmaX { get; set; }    // Biegespannung X
        public double SigmaY { get; set; }    // Biegespannung Y
        public double TauXY { get; set; }     // Schubspannung in XY-Ebene
        public double TauT { get; set; }      // Torsionsschubspannung
        public double Tau { get; set; }       // Gesamtschubspannung
        public double SigmaV { get; set; }    // Von-Mises-Spannung
    }
} 