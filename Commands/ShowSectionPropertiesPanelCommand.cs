using Rhino.Commands;
using Rhino.UI;
using Rhino;
using System;

namespace Moments_of_Inertia.Commands
{
    public class ShowSectionPropertiesPanelCommand : Command
    {
        public ShowSectionPropertiesPanelCommand()
        {
            Instance = this;
        }

        public static ShowSectionPropertiesPanelCommand Instance { get; private set; }

        public override string EnglishName => "ShowInertiaPropertiesPanel";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            var panelId = new Guid("284ae1ed-c3aa-45c4-9167-7934731c712f");
            Panels.OpenPanel(panelId);
            return Result.Success;
        }
    }
} 