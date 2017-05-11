using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class QuickFixProviderTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ProviderDoesNotKnowAboutInspection()
        {
            const string inputCode =
@"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var quickFixProvider = new QuickFixProvider(state, new IQuickFix[] {});
            Assert.AreEqual(0, quickFixProvider.QuickFixes(inspectionResults.First()).Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ProviderKnowsAboutInspection()
        {
            const string inputCode =
@"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var quickFixProvider = new QuickFixProvider(state, new IQuickFix[] {new RemoveUnusedDeclarationQuickFix(state)});
            Assert.AreEqual(1, quickFixProvider.QuickFixes(inspectionResults.First()).Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ResultDisablesFix()
        {
            const string inputCode =
@"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var quickFixProvider = new QuickFixProvider(state, new IQuickFix[] { new RemoveUnusedDeclarationQuickFix(state) });

            var result = inspectionResults.First();
            result.Properties.Add("DisableFixes", nameof(RemoveUnusedDeclarationQuickFix));

            Assert.AreEqual(0, quickFixProvider.QuickFixes(result).Count());
        }
    }
}