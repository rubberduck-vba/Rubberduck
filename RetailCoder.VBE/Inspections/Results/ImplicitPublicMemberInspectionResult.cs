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
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly QualifiedContext<ParserRuleContext> _qualifiedContext;

        public ImplicitPublicMemberInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext, Declaration item)
            : base(inspection, item)
        {
            _qualifiedContext = qualifiedContext;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new SpecifyExplicitPublicModifierQuickFix(_qualifiedContext.Context, QualifiedSelection), 
                    new IgnoreOnceQuickFix(_qualifiedContext.Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

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
