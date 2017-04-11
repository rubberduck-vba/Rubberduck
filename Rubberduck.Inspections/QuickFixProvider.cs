using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class QuickFixProvider : IQuickFixProvider
    {
        private readonly RubberduckParserState _state;
        private readonly Dictionary<Type, HashSet<IQuickFix>> _quickFixes = new Dictionary<Type, HashSet<IQuickFix>>();

        public QuickFixProvider(RubberduckParserState state, IEnumerable<IQuickFix> quickFixes)
        {
            _state = state;
            foreach (var quickFix in quickFixes)
            {
                foreach (var supportedInspection in quickFix.SupportedInspections)
                {
                    if (_quickFixes.ContainsKey(supportedInspection))
                    {
                        _quickFixes[supportedInspection].Add(quickFix);
                    }
                    else
                    {
                        _quickFixes.Add(supportedInspection, new HashSet<IQuickFix> {quickFix});
                    }
                }
            }
        }

        public IEnumerable<IQuickFix> QuickFixes(IInspectionResult result)
        {
            return _quickFixes[result.Inspection.GetType()];
        }

        private bool CanFix(IQuickFix fix, Type inspection)
        {
            return _quickFixes.ContainsKey(inspection) && _quickFixes[inspection].Contains(fix);
        }

        public void Fix(IQuickFix fix, IInspectionResult result)
        {
            if (!CanFix(fix, result.Inspection.GetType()))
            {
                throw new ArgumentException("Fix does not support this inspection.", nameof(result));
            }

            fix.Fix(result);
            _state.GetRewriter(result.QualifiedSelection.QualifiedName).Rewrite();
            _state.OnParseRequested(this);
        }

        public void FixInProcedure(IQuickFix fix, QualifiedSelection selection, Type inspectionType,
            IEnumerable<IInspectionResult> results)
        {
            if (!CanFix(fix, inspectionType))
            {
                throw new ArgumentException("Fix does not support this inspection.", nameof(inspectionType));
            }

            throw new NotImplementedException("A qualified selection does not state which proc we are in, so we could really only use this on inspection results that expose a Declaration.");
        }

        public void FixInModule(IQuickFix fix, QualifiedSelection selection, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            if (!CanFix(fix, inspectionType))
            {
                throw new ArgumentException("Fix does not support this inspection.", nameof(inspectionType));
            }

            var filteredResults = results
                .Where(result => result.Inspection.GetType() == inspectionType
                              && result.QualifiedSelection.QualifiedName == selection.QualifiedName)
                .ToList();

            foreach (var result in filteredResults)
            {
                fix.Fix(result);
            }

            if (filteredResults.Any())
            {
                _state.GetRewriter(filteredResults.First().QualifiedSelection.QualifiedName).Rewrite();
                _state.OnParseRequested(this);
            }
        }

        public void FixInProject(IQuickFix fix, QualifiedSelection selection, Type inspectionType,
            IEnumerable<IInspectionResult> results)
        {
            if (!CanFix(fix, inspectionType))
            {
                throw new ArgumentException("Fix does not support this inspection.", nameof(inspectionType));
            }

            var filteredResults = results
                .Where(result => result.Inspection.GetType() == inspectionType
                              && result.QualifiedSelection.QualifiedName.ProjectId == selection.QualifiedName.ProjectId)
                .ToList();

            foreach (var result in filteredResults)
            {
                fix.Fix(result);
            }

            if (filteredResults.Any())
            {
                var modules = filteredResults.Select(s => s.QualifiedSelection.QualifiedName).Distinct();
                foreach (var module in modules)
                {
                    _state.GetRewriter(module).Rewrite();
                }

                _state.OnParseRequested(this);
            }
        }

        public void FixAll(IQuickFix fix, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            if (!CanFix(fix, inspectionType))
            {
                throw new ArgumentException("Fix does not support this inspection.", nameof(inspectionType));
            }

            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType);

            foreach (var result in filteredResults)
            {
                fix.Fix(result);
            }

            if (filteredResults.Any())
            {
                var modules = filteredResults.Select(s => s.QualifiedSelection.QualifiedName).Distinct();
                foreach (var module in modules)
                {
                    _state.GetRewriter(module).Rewrite();
                }

                _state.OnParseRequested(this);
            }
        }

        public bool HasQuickFixes(IInspectionResult inspectionResult)
        {
            return _quickFixes[inspectionResult.Inspection.GetType()].Any();
        }
    }
}