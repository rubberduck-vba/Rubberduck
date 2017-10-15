using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteTypeHintInspectionResult : InspectionResultBase
    {
        private readonly string _result;
        private readonly Declaration _declaration;
        private IEnumerable<QuickFixBase> _quickFixes;

        public ObsoleteTypeHintInspectionResult(IInspection inspection, string result, QualifiedContext qualifiedContext, Declaration declaration)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
            _result = result;
            _declaration = declaration;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new RemoveTypeHintsQuickFix(Context, QualifiedSelection, _declaration), 
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return _result.Captialize(); }
        }
    }
}
