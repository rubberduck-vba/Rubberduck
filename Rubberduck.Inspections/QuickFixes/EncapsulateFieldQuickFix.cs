using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings.EncapsulateField;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class EncapsulateFieldQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(EncapsulatePublicFieldInspection)
        };

        public EncapsulateFieldQuickFix(RubberduckParserState state, IIndenter indenter)
        {
            _state = state;
            _indenter = indenter;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var vbe = result.Target.Project.VBE;

            using (var view = new EncapsulateFieldDialog(new EncapsulateFieldViewModel(_state, _indenter)))
            {
                var factory = new EncapsulateFieldPresenterFactory(vbe, _state, view);
                var refactoring = new EncapsulateFieldRefactoring(vbe, _indenter, factory);
                refactoring.Refactor(result.Target);
            }
        }

        public string Description(IInspectionResult result)
        {
            return string.Format(InspectionsUI.EncapsulatePublicFieldInspectionQuickFix, result.Target.IdentifierName);
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => false;
        public bool CanFixInProject => false;
    }
}