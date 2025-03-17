using Eto.Forms;
using Eto.Drawing;
using Rhino.UI;
using System;

namespace Moments_of_Inertia.Views
{
    [System.Runtime.InteropServices.Guid("284ae1ed-c3aa-45c4-9167-7934731c712f")]
    public class InertiaPropertiesPanel : Panel
    {
        private DropDown materialDropdown;
        private Button assignOutlineButton;
        private Button assignHollowsButton;
        private DropDown unitDropdown;
        private Button calculateButton;
        private Label resultLabel;

        public InertiaPropertiesPanel()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            // Create controls
            materialDropdown = new DropDown
            {
                Items = { "Steel", "Aluminum", "Wood" }
            };

            assignOutlineButton = new Button
            {
                Text = "Assign Outline"
            };
            assignOutlineButton.Click += AssignOutlineButton_Click;

            assignHollowsButton = new Button
            {
                Text = "Assign Hollows"
            };
            assignHollowsButton.Click += AssignHollowsButton_Click;

            unitDropdown = new DropDown
            {
                Items = { "mm", "cm" }
            };

            calculateButton = new Button
            {
                Text = "Calculate"
            };
            calculateButton.Click += CalculateButton_Click;

            resultLabel = new Label
            {
                Text = "Results will appear here"
            };

            // Layout
            var layout = new DynamicLayout { Padding = new Padding(10) };
            layout.AddRow("Material:", materialDropdown);
            layout.AddRow(assignOutlineButton);
            layout.AddRow(assignHollowsButton);
            layout.AddRow("Units:", unitDropdown);
            layout.AddRow(calculateButton);
            layout.AddRow(resultLabel);

            Content = layout;
        }

        private void AssignOutlineButton_Click(object sender, EventArgs e)
        {
            resultLabel.Text = "Outline button clicked";
        }

        private void AssignHollowsButton_Click(object sender, EventArgs e)
        {
            resultLabel.Text = "Hollows button clicked";
        }

        private void CalculateButton_Click(object sender, EventArgs e)
        {
            resultLabel.Text = "Calculate button clicked";
        }
    }
} 