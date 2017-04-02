using Rubberduck.UI;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameProjectHandler : RenameHandlerBase
    {
        public RenameProjectHandler(RenameModel model, IMessageBox messageBox)
            : base(model, messageBox) { }

        override public string ErrorMessage { get { return RubberduckUI.RenameDialog_ProjectRenameError; } }

        override public void Rename()
        {
            var projects = Model.VBE.VBProjects;
            var project = projects.SingleOrDefault(p => p.HelpFile == Model.Target.ProjectId);
            {
                if (project != null)
                {
                    project.Name = Model.NewName;
                }
            }
        }
    }
}
