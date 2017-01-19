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
    public sealed class ImplicitVariantReturnTypeInspectionResult : InspectionResultBase
    {
        private readonly string _identifierName;
        private IEnumerable<QuickFixBase> _quickFixes;

        public ImplicitVariantReturnTypeInspectionResult(IInspection inspection, string identifierName, QualifiedContext<ParserRuleContext> qualifiedContext, Declaration target)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context, target)
        {
            _identifierName = identifierName;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new SetExplicitVariantReturnTypeQuickFix(Context, QualifiedSelection, InspectionsUI.SetExplicitVariantReturnTypeQuickFix), 
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.ImplicitVariantReturnTypeInspectionResultFormat, _identifierName);
            }
        }

        public override NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Target);
        }
    }
}
