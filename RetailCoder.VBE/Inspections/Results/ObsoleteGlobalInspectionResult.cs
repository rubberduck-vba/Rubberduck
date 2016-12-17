using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteGlobalInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ObsoleteGlobalInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> context)
            : base(inspection, context.ModuleName, context.Context)
        {
            _quickFixes = new QuickFixBase[]
            {
                new ReplaceGlobalModifierQuickFix(Context, QualifiedSelection),
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.ObsoleteGlobalInspectionResultFormat, Target.DeclarationType.ToLocalizedString(), Target.IdentifierName).Captialize();
            }
        }
    }
}
