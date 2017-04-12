using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class QuickFixProviderTests
    {
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
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
            quickFixProvider.Fix(new OptionExplicitQuickFix(state), inspectionResults.First());
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
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
            quickFixProvider.Fix(new OptionExplicitQuickFix(state), inspectionResults.First());
        }
    }
}