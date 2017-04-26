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
            return _quickFixes.ContainsKey(result.Inspection.GetType())
                ? _quickFixes[result.Inspection.GetType()]
                : Enumerable.Empty<IQuickFix>();
        }

        private bool CanFix(IQuickFix fix, Type inspection)
        {
            return _quickFixes.ContainsKey(inspection) && _quickFixes[inspection].Contains(fix);
        }

        public void Fix(IQuickFix fix, IInspectionResult result)
        {
            if (!CanFix(fix, result.Inspection.GetType()))
            {
                throw new NotSupportedException($"{fix.GetType().Name} does not support this inspection.");
            }

            fix.Fix(result);
            _state.GetRewriter(result.QualifiedSelection.QualifiedName).Rewrite();
            _state.OnParseRequested(this);
        }

        public void FixInProcedure(IQuickFix fix, QualifiedMemberName? qualifiedMember, Type inspectionType,
            IEnumerable<IInspectionResult> results)
        {
            if (!CanFix(fix, inspectionType))
            {
                throw new NotSupportedException($"{fix.GetType().Name} does not support this inspection.");
            }

            Debug.Assert(qualifiedMember.HasValue, "Null qualified member.");

            var filteredResults = results
                .Where(result => result.Inspection.GetType() == inspectionType
                                 && result.QualifiedMemberName == qualifiedMember)
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

        public void FixInModule(IQuickFix fix, QualifiedSelection selection, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            if (!CanFix(fix, inspectionType))
            {
                throw new NotSupportedException($"{fix.GetType().Name} does not support this inspection.");
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
                throw new NotSupportedException($"{fix.GetType().Name} does not support this inspection.");
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
                throw new NotSupportedException($"{fix.GetType().Name} does not support this inspection.");
            }

            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType).ToArray();

            foreach(var result in filteredResults)
            {
                fix.Fix(result);
            }

            if (filteredResults.Any())
            {
                var modules = filteredResults.Select(s => s.QualifiedSelection.QualifiedName).Distinct();
                foreach (var module in modules)
                {
                    var moduleRewriter = _state.GetRewriter(module);
                    if (moduleRewriter.IsDirty)
                    {
                        _state.GetRewriter(module).Rewrite();
                        continue;
                    }

                    var attributesRewriter = _state.GetAttributeRewriter(module);
                    if (attributesRewriter.IsDirty)
                    {
                        _state.GetAttributeRewriter(module).Rewrite();
                        continue;
                    }
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