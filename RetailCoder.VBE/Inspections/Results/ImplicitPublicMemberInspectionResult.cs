using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitPublicMemberInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ImplicitPublicMemberInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext, Declaration item)
            : base(inspection, item)
        {
            _quickFixes = new QuickFixBase[]
            {
                new SpecifyExplicitPublicModifierQuickFix(qualifiedContext.Context, QualifiedSelection), 
                new IgnoreOnceQuickFix(qualifiedContext.Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.ImplicitPublicMemberInspectionResultFormat, Target.IdentifierName);
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
