using System;
using System.Collections.Generic;
using System.Threading;
using Antlr4.Runtime.Tree;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
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

        public IEnumerable<IInspectionResult> InspectionResultsForModules(IEnumerable<(string name, string content, ComponentType componentType)> modules, IDictionary<string,IEnumerable<string>> documentModuleSupertypeNames = null)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;
            return InspectionResults(vbe, documentModuleSupertypeNames);
        }

        public IEnumerable<IInspectionResult> InspectionResultsForModules((string name, string content, ComponentType componentType) module, params ReferenceLibrary[] libraries)
            => InspectionResultsForModules(new (string, string, ComponentType)[] { module }, libraries);

        public IEnumerable<IInspectionResult> InspectionResultsForModules(IEnumerable<(string name, string content, ComponentType componentType)> modules, ReferenceLibrary library, IDictionary<string, IEnumerable<string>> documentModuleSupertypeNames = null)
            => InspectionResultsForModules(modules, new ReferenceLibrary[] { library }, documentModuleSupertypeNames);

        public IEnumerable<IInspectionResult> InspectionResultsForModules(IEnumerable<(string name, string content, ComponentType componentType)> modules, IEnumerable<ReferenceLibrary> libraries, IDictionary<string, IEnumerable<string>> documentModuleSupertypeNames = null)
        {
            var vbe = MockVbeBuilder.BuildFromModules(modules, libraries).Object;
            return InspectionResults(vbe, documentModuleSupertypeNames);
        }

        public IEnumerable<IInspectionResult> InspectionResults(IVBE vbe, IDictionary<string, IEnumerable<string>> documentModuleSupertypeNames = null)
        {
            using (var state = MockParser.CreateAndParse(vbe, documentModuleSupertypeNames:documentModuleSupertypeNames))
            {
                var inspection = InspectionUnderTest(state);
                return InspectionResults(inspection, state);
            }
        }

        private static IEnumerable<IInspectionResult> InspectionResults(IInspection inspection, RubberduckParserState state)
        {
            if (inspection is IParseTreeInspection parseTreeInspection)
            {
                WalkTrees(parseTreeInspection, state);
            }

            return inspection.GetInspectionResults(CancellationToken.None);
        }

        protected static void WalkTrees(IParseTreeInspection inspection, RubberduckParserState state)
        {
            var codeKind = inspection.TargetKindOfCode;
            var listener = inspection.Listener;
            
            List<KeyValuePair<QualifiedModuleName, IParseTree>> trees;
            switch (codeKind)
            {
                case CodeKind.AttributesCode:
                    trees = state.AttributeParseTrees;
                    break;
                case CodeKind.CodePaneCode:
                    trees = state.ParseTrees;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(codeKind), codeKind, null);
            }

            foreach (var (module, tree) in trees)
            {
                listener.CurrentModuleName = module;
                ParseTreeWalker.Default.Walk(listener, tree);
            }
        }
    }
}