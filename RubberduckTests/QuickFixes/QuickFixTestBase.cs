using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Inspections;
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
            Func<RubberduckParserState, IInspection> inspectionFactory)
        {
            return ApplyQuickFixToAppropriateInspectionResults(
                inputCode,
                inspectionFactory,
                ApplyToFirstResult);
        }

        private string ApplyQuickFixToAppropriateInspectionResults(string inputCode,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Action<IQuickFix, IEnumerable<IInspectionResult>, IRewriteSession> applyQuickFix)
        {
            var vbe = TestVbe(inputCode, out var component);
            var (state, rewriteManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var inspection = inspectionFactory(state);
                var inspectionResults = InspectionResults(inspection, state);
                var rewriteSession = rewriteManager.CheckOutCodePaneSession();

                var quickFix = QuickFix(state);

                applyQuickFix(quickFix, inspectionResults, rewriteSession);

                return rewriteSession.CheckOutModuleRewriter(component.QualifiedModuleName).GetText();
            }
        }

        private IEnumerable<IInspectionResult> InspectionResults(IInspection inspection, RubberduckParserState state)
        {
            if (inspection is IParseTreeInspection)
            {
                var inspector = InspectionsHelper.GetInspector(inspection);
                return inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            }

            return inspection.GetInspectionResults(CancellationToken.None);
        }

        private void ApplyToFirstResult(IQuickFix quickFix, IEnumerable<IInspectionResult> inspectionResults, IRewriteSession rewriteSession)
        {
            var resultToFix = inspectionResults.First();
            quickFix.Fix(resultToFix, rewriteSession);
        }

        protected string ApplyQuickFixToAllInspectionResults(string inputCode,
            Func<RubberduckParserState, IInspection> inspectionFactory)
        {
            return ApplyQuickFixToAppropriateInspectionResults(
                inputCode,
                inspectionFactory,
                ApplyToAllResults);
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
        Func<RubberduckParserState, IInspection> inspectionFactory)
        {
            return ApplyQuickFixToAppropriateInspectionResultsForImplementedInterface(
                interfaceInputCode,
                implementationInputCode,
                inspectionFactory,
                ApplyToFirstResult);
        }

        private (string interfaceCode, string implementationCode) ApplyQuickFixToAppropriateInspectionResultsForImplementedInterface(
            string interfaceCode, 
            string implementationCode,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Action<IQuickFix, IEnumerable<IInspectionResult>, IRewriteSession> applyQuickFix)
        {
            var (vbe, interfaceModuleName, implementationModuleName) = TestVbeForImplementedInterface(interfaceCode, implementationCode);

            var (state, rewriteManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var inspection = inspectionFactory(state);
                var inspectionResults = InspectionResults(inspection, state);
                var rewriteSession = rewriteManager.CheckOutCodePaneSession();

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
            Func<IInspectionResult, bool> predicate)
        {
            var applyQuickFix = ApplyToFirstResultSatisfyingPredicateAction(predicate);
            return ApplyQuickFixToAppropriateInspectionResultsForImplementedInterface(
                interfaceInputCode,
                implementationInputCode,
                inspectionFactory,
                applyQuickFix);
        }

        private Action<IQuickFix, IEnumerable<IInspectionResult>, IRewriteSession> ApplyToFirstResultSatisfyingPredicateAction(Func<IInspectionResult, bool> predicate)
        {
            return (quickFix, inspectionResults, rewriteSession) =>
                quickFix.Fix(inspectionResults.First(predicate), rewriteSession);
        }

        protected string ApplyQuickFixToFirstInspectionResultSatisfyingPredicate(
            string code,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Func<IInspectionResult, bool> predicate)
        {
            var applyQuickFix = ApplyToFirstResultSatisfyingPredicateAction(predicate);
            return ApplyQuickFixToAppropriateInspectionResults(
                code,
                inspectionFactory,
                applyQuickFix);
        }
    }
}