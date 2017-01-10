using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    class AggregateInspectionResult: InspectionResultBase
    {
        private readonly List<IInspectionResult> _results;
        private readonly IEnumerable<QuickFixBase> _quickFixes;
        private readonly IInspectionResult _result;

        public AggregateInspectionResult(List<IInspectionResult> results, IEnumerable<QuickFixBase> quickFixes = null)
            : base(results[0].Inspection, results[0].QualifiedSelection.QualifiedName, ParserRuleContext.EmptyContext)
        {
            _results = results;
            _result = results[0];
            _quickFixes = quickFixes ?? Enumerable.Empty<QuickFixBase>();
        }
        
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

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override int CompareTo(IInspectionResult other)
        {
            if (other == this)
            {
                return 0;
            }
            var result = other as AggregateInspectionResult;
            if (result == null)
            {
                return -1;
            }
            if (_results.Count != result._results.Count) {
                return _results.Count - result._results.Count;
            }
            for (var i = 0; i < _results.Count; i++)
            {
                if (_results[i].CompareTo(result._results[i]) != 0)
                {
                    return _results[i].CompareTo(result._results[i]);
                }
            }
            return 0;
        }
    }
}
