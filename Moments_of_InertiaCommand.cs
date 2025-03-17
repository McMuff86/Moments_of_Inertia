using System;
using Rhino;
using Rhino.Commands;
using Rhino.UI;

namespace Moments_of_Inertia
{
    public class Moments_of_InertiaCommand : Command
    {
        public Moments_of_InertiaCommand()
        {
            // Rhino only creates one instance of each command class defined in a
            // plug-in, so it is safe to store a refence in a static property.
            Instance = this;
        }

        ///<summary>The only instance of this command.</summary>
        public static Moments_of_InertiaCommand Instance { get; private set; }

        ///<returns>The command name as it appears on the Rhino command line.</returns>
        public override string EnglishName => "MomentsOfInertia";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // Open the InertiaPropertiesPanel
            var panelId = new Guid("284ae1ed-c3aa-45c4-9167-7934731c712f");
            Panels.OpenPanel(panelId);
            
            return Result.Success;
        }
    }
}
