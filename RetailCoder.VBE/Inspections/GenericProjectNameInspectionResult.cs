using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class GenericProjectNameInspectionResult : CodeInspectionResultBase
    {
        public GenericProjectNameInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedModuleName qualifiedName) 
            : base(inspection, type, new CommentNode("", new QualifiedSelection(qualifiedName, Selection.Home)))
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>();
            /*{
                { "Rename Project", RenameProject }
            };*/
        }

        private void RenameProject()
        {
            /*var progress = new ParsingProgressPresenter();
            var result = progress.Parse(_parser, IDE.ActiveVBProject);

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(IDE, view, result);
                var refactoring = new RenameRefactoring(factory);
                refactoring.Refactor(target);
            }*/
        }
    }
}
