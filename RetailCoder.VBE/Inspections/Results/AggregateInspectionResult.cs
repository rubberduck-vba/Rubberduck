using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;

namespace Rubberduck.Inspections.Results
{
    public class AggregateInspectionResult: InspectionResultBase
    {
        private readonly IInspectionResult _result;
        private readonly int _count;

        public AggregateInspectionResult(IInspectionResult firstResult, int count)
            : base(firstResult.Inspection, firstResult.QualifiedSelection.QualifiedName, ParserRuleContext.EmptyContext)
        {
            _result = firstResult;
            _count = count;
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.AggregateInspectionResultFormat, _result.Inspection.Description, _count);
            }
        }

        public override QualifiedSelection QualifiedSelection { get { return _result.QualifiedSelection; } }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get { return _result.QuickFixes == null ? base.QuickFixes : new[] { _result.QuickFixes.FirstOrDefault() }; }
        }

        public override QuickFixBase DefaultQuickFix { get { return _result.QuickFixes == null ? null : _result.QuickFixes.FirstOrDefault(); } }

        public override int CompareTo(IInspectionResult other)
        {
            if (other == this)
            {
                return 0;
            }
            var aggregated = other as AggregateInspectionResult;
            if (aggregated == null)
            {
                return -1;
            }
            if (_count != aggregated._count) {
                return _count - aggregated._count;
            }
            for (var i = 0; i < _count; i++)
            {
                if (_result.CompareTo(aggregated._result) != 0)
                {
                    return _result.CompareTo(aggregated._result);
                }
            }
            return 0;
        }
    }
}
