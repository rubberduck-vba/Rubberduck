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
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public IdentifierNotUsedInspectionResult(IInspection inspection, Declaration target,
            ParserRuleContext context, QualifiedModuleName qualifiedName)
            : base(inspection, qualifiedName, context, target)
        {
            _quickFixes = new QuickFixBase[]
            {
                new RemoveUnusedDeclarationQuickFix(context, QualifiedSelection, Target), 
                new IgnoreOnceQuickFix(context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }
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
