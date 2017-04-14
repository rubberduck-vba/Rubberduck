using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings.Rename;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RenameDeclarationQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(HungarianNotationInspection),
            typeof(UseMeaningfulNameInspection),
            typeof(DefaultProjectNameInspection)
        };

        public RenameDeclarationQuickFix(RubberduckParserState state, IMessageBox messageBox)
        {
            _state = state;
            _messageBox = messageBox;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var vbe = result.Target.Project.VBE;

            using (var view = new RenameDialog(new RenameViewModel(_state)))
            {
                var factory = new RenamePresenterFactory(vbe, view, _state);
                var refactoring = new RenameRefactoring(vbe, factory, _messageBox, _state);
                refactoring.Refactor(result.Target);
            }
        }

        public string Description(IInspectionResult result)
        {
            return string.Format(RubberduckUI.Rename_DeclarationType,
                RubberduckUI.ResourceManager.GetString("DeclarationType_" + result.Target.DeclarationType,
                    CultureInfo.CurrentUICulture));
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => false;
        public bool CanFixInProject => false;
    }
}