using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Results
{
    class AggregateInspectionResult: IInspectionResult
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;
        private readonly List<IInspectionResult> _results;

        public AggregateInspectionResult(List<IInspectionResult> encapsulatedResults)
        {
            if (encapsulatedResults.Count == 0)
            {
                throw new InvalidOperationException("Cannot aggregate results without results, wtf?");
            }
            _results = encapsulatedResults;
            _quickFixes = new QuickFixBase[]{};
            // If there were any reasonable way to provide this mess with a fix, we'd put it here
        }

        //public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public string Description
        {
            get
            {
                return string.Format(InspectionsUI.AggregateInspectionResultFormat, _results[0].Inspection.Meta, _results.Count);
            }
        }

        public IInspection Inspection
        {
            get
            {
                return _results[0].Inspection;
            }
        }

        public QualifiedSelection QualifiedSelection
        {
            get
            {
                return _results[0].QualifiedSelection;
            }
        }

        public IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public int CompareTo(object obj)
        {
            throw new NotImplementedException();
        }

        public int CompareTo(IInspectionResult other)
        {
            throw new NotImplementedException();
        }

        public object[] ToArray()
        {
            throw new NotImplementedException();
        }

        
    }
}
