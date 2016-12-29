using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Results
{
    class AggregateInspectionResult: InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;
        private readonly List<IInspectionResult> _results;

        public AggregateInspectionResult(List<IInspectionResult> encapsulatedResults) 
            : base(null, new QualifiedModuleName(), null)
        {
            if (encapsulatedResults.Count == 0)
            {
                throw new InvalidOperationException("Cannot aggregate results without results, wtf?");
            }
            _results = encapsulatedResults;
            _quickFixes = new QuickFixBase[]{};
            // If there were any reasonable way to provide this mess with a fix, we'd put it here
        }

        public override string Description
        {
            get
            {
                return string.Format(InspectionsUI.AggregateInspectionResultFormat, _results[0].Inspection.Meta, _results.Count);
            }
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }
    }
}
