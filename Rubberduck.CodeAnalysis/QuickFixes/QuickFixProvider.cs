﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.CSharp.RuntimeBinder;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
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
                    try
                    {
                        value = result.Properties.DisableFixes;
                    }
                    catch (RuntimeBinderException)
                    {
                        return true;
                    }

                    if (value == null)
                    {
                        return true;
                    }

                    return !value.Split(',').Contains(fix.GetType().Name);
                })
                .OrderBy(fix => fix.SupportedInspections.Count); // most specific fixes first; keeps "ignore once" last
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
            _state.RewriteAllModules();
            _state.OnParseRequested(this);
        }

        public void FixInProcedure(IQuickFix fix, QualifiedMemberName? qualifiedMember, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            Debug.Assert(qualifiedMember.HasValue, "Null qualified member.");

            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType && result.QualifiedMemberName == qualifiedMember).ToList();

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
                _state.RewriteAllModules();
                _state.OnParseRequested(this);
            }
        }

        public void FixInModule(IQuickFix fix, QualifiedSelection selection, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType && result.QualifiedSelection.QualifiedName == selection.QualifiedName).ToList();

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
                _state.RewriteAllModules();
                _state.OnParseRequested(this);
            }
        }

        public void FixInProject(IQuickFix fix, QualifiedSelection selection, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType && result.QualifiedSelection.QualifiedName.ProjectId == selection.QualifiedName.ProjectId).ToList();

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
                _state.RewriteAllModules();
                _state.OnParseRequested(this);
            }
        }

        public void FixAll(IQuickFix fix, Type inspectionType, IEnumerable<IInspectionResult> results)
        {
            var filteredResults = results.Where(result => result.Inspection.GetType() == inspectionType).ToArray();

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
                _state.RewriteAllModules();
                _state.OnParseRequested(this);
            }
        }

        public bool HasQuickFixes(IInspectionResult inspectionResult)
        {
            return _quickFixes.ContainsKey(inspectionResult.Inspection.GetType()) && _quickFixes[inspectionResult.Inspection.GetType()].Any();
        }
    }
}