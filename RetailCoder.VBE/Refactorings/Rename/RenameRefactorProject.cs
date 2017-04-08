﻿using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactorProject : RenameRefactorBase
    {
        private readonly RenameModel _model;

        public RenameRefactorProject(RenameModel model)
            : base(model)
        {
            _model = model;
        }


        public override string ErrorMessage { get { return RubberduckUI.RenameDialog_ProjectRenameError; } }

        public override bool RequestParseAfterRename { get { return false; } }

        public override void Rename(Declaration renameTarget, string newName)
        {
            var projects = _model.VBE.VBProjects;
            var project = projects.SingleOrDefault(p => p.HelpFile == _model.Target.ProjectId);
            {
                if (project != null)
                {
                    project.Name = newName;
                }
            }
        }
    }
}
