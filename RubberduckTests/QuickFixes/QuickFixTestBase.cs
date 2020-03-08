using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Antlr4.Runtime.Tree;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    public abstract class QuickFixTestBase
    {
        protected abstract IQuickFix QuickFix(RubberduckParserState state);

        protected virtual IVBE TestVbe(string code, out IVBComponent component)
        {
            return MockVbeBuilder.BuildFromSingleStandardModule(code, out component).Object;
        }

        protected string ApplyQuickFixToFirstInspectionResult(string inputCode,
            Func<RubberduckParserState, IInspection> inspectionFactory, 
            CodeKind codeKind = CodeKind.CodePaneCode)
        {
            return ApplyQuickFixToAppropriateInspectionResults(
                inputCode,
                inspectionFactory,
                ApplyToFirstResult,
                codeKind);
        }

        protected string ApplyQuickFixToFirstInspectionResult(
            IVBE vbe,
            string componentName,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            CodeKind codeKind = CodeKind.CodePaneCode)
        {
            return ApplyQuickFixToAppropriateInspectionResults(
                vbe,
                componentName,
                inspectionFactory,
                ApplyToFirstResult,
                codeKind);
        }

        private string ApplyQuickFixToAppropriateInspectionResults(string inputCode,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Action<IQuickFix, IEnumerable<IInspectionResult>, IRewriteSession> applyQuickFix,
            CodeKind codeKind)
        {
            var vbe = TestVbe(inputCode, out var component);
            return ApplyQuickFixToAppropriateInspectionResults(
                vbe,
                component.Name,
                inspectionFactory,
                applyQuickFix,
                codeKind);
        }

        private string ApplyQuickFixToAppropriateInspectionResults(
            IVBE vbe,
            string componentName,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Action<IQuickFix, IEnumerable<IInspectionResult>, IRewriteSession> applyQuickFix,
            CodeKind codeKind)
        {
            var (state, rewriteManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var inspection = inspectionFactory(state);
                var inspectionResults = InspectionResults(inspection, state);
                var rewriteSession = codeKind == CodeKind.AttributesCode
                    ? rewriteManager.CheckOutAttributesSession()
                    : rewriteManager.CheckOutCodePaneSession();

                var quickFix = QuickFix(state);

                applyQuickFix(quickFix, inspectionResults, rewriteSession);

                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName == componentName);

                return rewriteSession.CheckOutModuleRewriter(module).GetText();
            }
        }

        private IEnumerable<IInspectionResult> InspectionResults(IInspection inspection, RubberduckParserState state)
        {
            if (inspection is IParseTreeInspection parseTreeInspection)
            {
                WalkTrees(parseTreeInspection, state);
            }

            return inspection.GetInspectionResults(CancellationToken.None);
        }

        private static void WalkTrees(IParseTreeInspection inspection, RubberduckParserState state)
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

        private void ApplyToFirstResult(IQuickFix quickFix, IEnumerable<IInspectionResult> inspectionResults, IRewriteSession rewriteSession)
        {
            var resultToFix = inspectionResults.First();
            quickFix.Fix(resultToFix, rewriteSession);
        }

        protected string ApplyQuickFixToAllInspectionResults(
            IVBE vbe,
            string componentName,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            CodeKind codeKind = CodeKind.CodePaneCode)
        {
            return ApplyQuickFixToAppropriateInspectionResults(
                vbe,
                componentName,
                inspectionFactory,
                ApplyToAllResults,
                codeKind);
        }

        protected string ApplyQuickFixToAllInspectionResults(string inputCode,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            CodeKind codeKind = CodeKind.CodePaneCode)
        {
            return ApplyQuickFixToAppropriateInspectionResults(
                inputCode,
                inspectionFactory,
                ApplyToAllResults,
                codeKind);
        }

        private void ApplyToAllResults(IQuickFix quickFix, IEnumerable<IInspectionResult> inspectionResults, IRewriteSession rewriteSession)
        {
            foreach (var inspectionResult in inspectionResults)
            {
                quickFix.Fix(inspectionResult, rewriteSession);
            }
        }



        protected (string interfaceCode, string implementationCode) ApplyQuickFixToFirstInspectionResultForImplementedInterface(
            string interfaceInputCode,
            string implementationInputCode,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            CodeKind codeKind = CodeKind.CodePaneCode)
        {
            return ApplyQuickFixToAppropriateInspectionResultsForImplementedInterface(
                interfaceInputCode,
                implementationInputCode,
                inspectionFactory,
                ApplyToFirstResult,
                codeKind);
        }

        private (string interfaceCode, string implementationCode) ApplyQuickFixToAppropriateInspectionResultsForImplementedInterface(
            string interfaceCode, 
            string implementationCode,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Action<IQuickFix, IEnumerable<IInspectionResult>, IRewriteSession> applyQuickFix,
            CodeKind codeKind)
        {
            var (vbe, interfaceModuleName, implementationModuleName) = TestVbeForImplementedInterface(interfaceCode, implementationCode);

            var (state, rewriteManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var inspection = inspectionFactory(state);
                var inspectionResults = InspectionResults(inspection, state);
                var rewriteSession = codeKind == CodeKind.AttributesCode 
                    ? rewriteManager.CheckOutAttributesSession() 
                    : rewriteManager.CheckOutCodePaneSession();

                var quickFix = QuickFix(state);

                applyQuickFix(quickFix, inspectionResults, rewriteSession);

                var actualInterfaceCode = rewriteSession.CheckOutModuleRewriter(interfaceModuleName).GetText();
                var actualImplementationCode = rewriteSession.CheckOutModuleRewriter(implementationModuleName).GetText();

                return (actualInterfaceCode, actualImplementationCode);
            }
        }

        protected virtual (IVBE vbe, QualifiedModuleName interfaceModuleName, QualifiedModuleName implemetingModuleName) TestVbeForImplementedInterface(string interfaceCode, string implementationCode)
        {
            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, interfaceCode)
                .AddComponent("Class1", ComponentType.ClassModule, implementationCode)
                .AddProjectToVbeBuilder()
                .Build().Object;

            var project = vbe.VBProjects[0];
            var interfaceComponent = project.VBComponents[0];
            var implementationComponent = project.VBComponents[1];

            return (vbe, interfaceComponent.QualifiedModuleName, implementationComponent.QualifiedModuleName);
        }

        protected (string interfaceCode, string implementationCode) ApplyQuickFixToFirstInspectionResultForImplementedInterfaceSatisfyingPredicate(
            string interfaceInputCode,
            string implementationInputCode,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Func<IInspectionResult, bool> predicate,
            CodeKind codeKind = CodeKind.CodePaneCode)
        {
            var applyQuickFix = ApplyToFirstResultSatisfyingPredicateAction(predicate);
            return ApplyQuickFixToAppropriateInspectionResultsForImplementedInterface(
                interfaceInputCode,
                implementationInputCode,
                inspectionFactory,
                applyQuickFix,
                codeKind);
        }

        private Action<IQuickFix, IEnumerable<IInspectionResult>, IRewriteSession> ApplyToFirstResultSatisfyingPredicateAction(Func<IInspectionResult, bool> predicate)
        {
            return (quickFix, inspectionResults, rewriteSession) =>
                quickFix.Fix(inspectionResults.First(predicate), rewriteSession);
        }

        protected string ApplyQuickFixToFirstInspectionResultSatisfyingPredicate(
            string code,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Func<IInspectionResult, bool> predicate,
            CodeKind codeKind = CodeKind.CodePaneCode)
        {
            var applyQuickFix = ApplyToFirstResultSatisfyingPredicateAction(predicate);
            return ApplyQuickFixToAppropriateInspectionResults(
                code,
                inspectionFactory,
                applyQuickFix,
                codeKind);
        }
    }
}