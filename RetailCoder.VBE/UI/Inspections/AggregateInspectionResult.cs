﻿using System.Collections.Generic;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Inspections
{
    public class AggregateInspectionResult: IInspectionResult
    {
        private readonly IInspectionResult _result;
        private readonly int _count;

        public AggregateInspectionResult(IInspectionResult firstResult, int count)
        {
            _result = firstResult;
            _count = count;
        }

        public string Description => string.Format(InspectionsUI.AggregateInspectionResultFormat, _result.Inspection.Description, _count);

        public QualifiedSelection QualifiedSelection => _result.QualifiedSelection;
        public IInspection Inspection => _result.Inspection;

        public Declaration Target => _result.Target;

        public IEnumerable<IQuickFix> QuickFixes => _result.QuickFixes;

        public bool HasQuickFixes => _result.HasQuickFixes;

        public IQuickFix DefaultQuickFix => _result.DefaultQuickFix;

        public int CompareTo(IInspectionResult other)
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

        public int CompareTo(object obj)
        {
            return CompareTo(obj as IInspectionResult);
        }
    }
}
