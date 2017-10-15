using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class IdentifierNotUsedInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly ParserRuleContext _context;

        public IdentifierNotUsedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, qualifiedName, context, target)
        {
            _context = context;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new RemoveUnusedDeclarationQuickFix(_context, QualifiedSelection, Target), 
                    new IgnoreOnceQuickFix(_context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description 
        {
            get
            {
                return string.Format(InspectionsUI.IdentifierNotUsedInspectionResultFormat, Target.DeclarationType.ToLocalizedString(), Target.IdentifierName).Captialize();
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
