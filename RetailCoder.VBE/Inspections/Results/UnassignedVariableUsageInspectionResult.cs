using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class UnassignedVariableUsageInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;

        public UnassignedVariableUsageInspectionResult(IInspection inspection, ParserRuleContext context, QualifiedModuleName qualifiedName, Declaration declaration)
            : base(inspection, qualifiedName, context, declaration)
        { }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    //new RemoveUnassignedVariableUsageQuickFix(Context, QualifiedSelection),   // removed until we can reinstate this for a specific reference
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.UnassignedVariableUsageInspectionResultFormat, Target.IdentifierName).Captialize();
            }
        }
    }
}
