using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;

namespace Rubberduck.Inspections.Results
{
    class AggregateInspectionResult: InspectionResultBase
    {
        private readonly List<IInspectionResult> _results;
        private readonly IInspectionResult _result;

        public AggregateInspectionResult(List<IInspectionResult> results)
            : base(results[0].Inspection, results[0].QualifiedSelection.QualifiedName, ParserRuleContext.EmptyContext)
        {
            _results = results;
            _result = results[0];
        }

        public IReadOnlyList<IInspectionResult> IndividualResults { get { return _results; } }
        
        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.AggregateInspectionResultFormat, _result.Inspection.Description, _results.Count);
            }
        }

        public override QualifiedSelection QualifiedSelection
        {
            get
            {
                return _result.QualifiedSelection;
            }
        }

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
            if (_results.Count != aggregated._results.Count) {
                return _results.Count - aggregated._results.Count;
            }
            for (var i = 0; i < _results.Count; i++)
            {
                if (_results[i].CompareTo(aggregated._results[i]) != 0)
                {
                    return _results[i].CompareTo(aggregated._results[i]);
                }
            }
            return 0;
        }
    }
}
