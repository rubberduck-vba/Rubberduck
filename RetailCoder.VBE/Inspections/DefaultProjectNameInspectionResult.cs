using System;
using System.Collections.Generic;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public class DefaultProjectNameInspectionResult : CodeInspectionResultBase
    {
        private readonly VBProjectParseResult _parseResult;
        private readonly ICodePaneWrapperFactory _wrapperFactory;

        public DefaultProjectNameInspectionResult(string inspection, CodeInspectionSeverity type, Declaration target, VBProjectParseResult parseResult, ICodePaneWrapperFactory wrapperFactory) 
            : base(inspection, type, target)
        {
            _parseResult = parseResult;
            _wrapperFactory = wrapperFactory;
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            var project = RubberduckUI.ResourceManager.GetString("DeclarationType_" + DeclarationType.Project, RubberduckUI.Culture);
            return new Dictionary<string, Action>
            {
                { string.Format(RubberduckUI.Rename_DeclarationType, project), RenameProject }
            };
        }

        private void RenameProject()
        {
            var vbe = QualifiedSelection.QualifiedName.Project.VBE;

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(vbe, view, _parseResult, new RubberduckMessageBox(), _wrapperFactory);
                var refactoring = new RenameRefactoring(factory, new ActiveCodePaneEditor(vbe, _wrapperFactory), new RubberduckMessageBox());
                refactoring.Refactor(Target);
            }
        }
    }
}
