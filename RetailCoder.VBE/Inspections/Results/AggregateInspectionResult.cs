using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Results
{
    class AggregateInspectionResult: IInspectionResult
    {
        private readonly List<IInspectionResult> _results;

        public AggregateInspectionResult(List<IInspectionResult> encapsulatedResults)
        {
            if (encapsulatedResults.Count == 0)
            {
                throw new InvalidOperationException("Cannot aggregate results without results, wtf?");
            }
            _results = encapsulatedResults;
        }
        

        public string Description
        {
            get
            {
                return string.Format(InspectionsUI.AggregateInspectionResultFormat, _results[0].Inspection.Description, _results.Count);
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
                return Enumerable.Empty<QuickFixBase>();
            }
        }

        public int CompareTo(object obj)
        {
            if (obj == this)
            {
                return 0;
            }
            IInspectionResult result = obj as IInspectionResult;
            return result == null ? -1 : CompareTo(result);
        }

        public int CompareTo(IInspectionResult other)
        {
            if (other == this)
            {
                return 0;
            }
            AggregateInspectionResult result = other as AggregateInspectionResult;
            if (result == null)
            {
                return -1;
            }
            if (_results.Count != result._results.Count) {
                return _results.Count - result._results.Count;
            }
            for (int i = 0; i < _results.Count; i++)
            {
                if (_results[i].CompareTo(result._results[i]) != 0)
                {
                    return _results[i].CompareTo(result._results[i]);
                }
            }
            return 0;
        }

        public object[] ToArray()
        {
            var module = QualifiedSelection.QualifiedName;
            return new object[] { Inspection.Severity.ToString(), module.ProjectName, module.ComponentName, Description, QualifiedSelection.Selection.StartLine, QualifiedSelection.Selection.StartColumn };
        }
    }
}
