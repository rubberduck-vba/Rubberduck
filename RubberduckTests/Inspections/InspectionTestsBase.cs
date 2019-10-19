using System;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    public abstract class InspectionTestsBase
    {
        protected abstract IInspection InspectionUnderTest(RubberduckParserState state);

        public IEnumerable<IInspectionResult> InspectionResultsForStandardModule(string code)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _).Object;
            return InspectionResults(vbe);
        }

        public IEnumerable<IInspectionResult> InspectionResultsForModules(params (string name, string content, ComponentType componentType)[] modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return InspectionResults(vbe);
        }

        public IEnumerable<IInspectionResult> InspectionResultsForModules(IEnumerable<(string name, string content, ComponentType componentType)> modules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return InspectionResults(vbe);
        }

        public IEnumerable<IInspectionResult> InspectionResultsForModules((string name, string content, ComponentType componentType) module, params string[] libraries)
            => InspectionResultsForModules(new (string, string, ComponentType)[] { module }, libraries);

        public IEnumerable<IInspectionResult> InspectionResultsForModules(IEnumerable<(string name, string content, ComponentType componentType)> modules, string library)
            => InspectionResultsForModules(modules, new string[] { library });

        public IEnumerable<IInspectionResult> InspectionResultsForModules(IEnumerable<(string name, string content, ComponentType componentType)> modules, IEnumerable<string> libraries)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules, libraries).Object;
            return InspectionResults(vbe);
        }

        public IEnumerable<IInspectionResult> InspectionResults(IVBE vbe)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var inspection = InspectionUnderTest(state);
                return InspectionResults(inspection, state);
            }
        }

        private static IEnumerable<IInspectionResult> InspectionResults(IInspection inspection, RubberduckParserState state)
        {
            if (inspection is IParseTreeInspection)
            {
                var inspector = InspectionsHelper.GetInspector(inspection);
                return inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            }

            return inspection.GetInspectionResults(CancellationToken.None);
        }
    }
}