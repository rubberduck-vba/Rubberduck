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

        //provide ability for inspection tests to modify the default
        //Project and Module names (e.g., the MockVbeBuilder defaults are flagged by MeaninglessName inspection)
        public string TestProjectName { set; get; } = MockVbeBuilder.TestProjectName;
        public string TestModuleName { set; get; } = MockVbeBuilder.TestModuleName;

        public IEnumerable<IInspectionResult> InspectionResultsForStandardModule(string code)
        {
            MockVbeBuilder.TestProjectName = TestProjectName;
            MockVbeBuilder.TestModuleName = TestModuleName;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _).Object;
            return InspectionResults(vbe);
        }

        public IEnumerable<IInspectionResult> InspectionResultsForModules(params (string name, string content, ComponentType componentType)[] modules)
        {
            MockVbeBuilder.TestProjectName = TestProjectName;
            MockVbeBuilder.TestModuleName = TestModuleName;
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return InspectionResults(vbe);
        }

        public IEnumerable<IInspectionResult> InspectionResultsForModules(IEnumerable<(string name, string content, ComponentType componentType)> modules)
        {
            MockVbeBuilder.TestProjectName = TestProjectName;
            MockVbeBuilder.TestModuleName = TestModuleName;
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return InspectionResults(vbe);
        }

        public IEnumerable<IInspectionResult> InspectionResultsForModules((string name, string content, ComponentType componentType) module, string library)
            => InspectionResultsForModules(new (string, string, ComponentType)[] { module }, new string[] { library });

        public IEnumerable<IInspectionResult> InspectionResultsForModules(IEnumerable<(string name, string content, ComponentType componentType)> modules, string library)
            => InspectionResultsForModules(modules, new string[] { library });

        public IEnumerable<IInspectionResult> InspectionResultsForModules((string name, string content, ComponentType componentType) module, IEnumerable<string> libraries)
            => InspectionResultsForModules(new(string, string, ComponentType)[] { module }, libraries);

        public IEnumerable<IInspectionResult> InspectionResultsForModules(IEnumerable<(string name, string content, ComponentType componentType)> modules, IEnumerable<string> libraries)
        {
            MockVbeBuilder.TestProjectName = TestProjectName;
            MockVbeBuilder.TestModuleName = TestModuleName;
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