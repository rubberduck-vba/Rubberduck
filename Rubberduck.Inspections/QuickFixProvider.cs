using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            if (!_quickFixes.ContainsKey(result.Inspection.GetType()))
            {
                return Enumerable.Empty<IQuickFix>();
            }

            return _quickFixes[result.Inspection.GetType()].Where(fix =>
            {
                string value;
                if (!result.Properties.TryGetValue("DisableFixes", out value)) { return true; }

                if (value.Split(',').Contains(fix.GetType().Name))
                {
                    return false;
                }

                return true;
            });
        }

        private bool CanFix(IQuickFix fix, IInspectionResult result)
        {
            return QuickFixes(result).Contains(fix);
        }

        public void Fix(IQuickFix fix, IInspectionResult result)
        {
            if (!CanFix(fix, result))
            {
                return;
            }

            fix.Fix(result);
            _state.GetRewriter(result.QualifiedSelection.QualifiedName).Rewrite();
            _state.OnParseRequested(this);
        }

        public void FixInProcedure(IQuickFix fix, QualifiedMemberName? qualifiedMember, Type inspectionType,
            IEnumerable<IInspectionResult> results)
        {
            Debug.Assert(qualifiedMember.HasValue, "Null qualified member.");

            var filteredResults = results
                .Where(result => result.Inspection.GetType() == inspectionType
                                 && result.QualifiedMemberName == qualifiedMember)
                .ToList();

            foreach (var result in filteredResults)
            {
                if (!CanFix(fix, result))
                {
                    continue;
                }

                fix.Fix(result);
            }

            if (filteredResults.Any())
            {
                _state.GetRewriter(filteredResults.First().QualifiedSelection.QualifiedName).Rewrite();
                _state.OnParseRequested(this);
            }
        }

        public void FixInModule(IQuickFix fix, QualifiedSelection selection, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            var filteredResults = results
                .Where(result => result.Inspection.GetType() == inspectionType
                              && result.QualifiedSelection.QualifiedName == selection.QualifiedName)
                .ToList();

            foreach (var result in filteredResults)
            {
                if (!CanFix(fix, result))
                {
                    continue;
                }

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
            var filteredResults = results
                .Where(result => result.Inspection.GetType() == inspectionType
                              && result.QualifiedSelection.QualifiedName.ProjectId == selection.QualifiedName.ProjectId)
                .ToList();

            foreach (var result in filteredResults)
            {
                if (!CanFix(fix, result))
                {
                    continue;
                }

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
            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType);

            foreach (var result in filteredResults)
            {
                if (!CanFix(fix, result))
                {
                    continue;
                }

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
            return _quickFixes.ContainsKey(inspectionResult.Inspection.GetType()) &&
                   _quickFixes[inspectionResult.Inspection.GetType()].Any();
        }
    }
}