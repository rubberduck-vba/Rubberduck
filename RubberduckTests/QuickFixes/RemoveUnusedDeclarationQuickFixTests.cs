using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class RemoveUnusedDeclarationQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ConstantNotUsed_QuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo()
Const const1 As Integer = 9
End Sub";

            const string expectedCode =
@"Public Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ConstantNotUsedInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new RemoveUnusedDeclarationQuickFix(state).Fix(inspectionResults.First());

            var rewriter = state.GetRewriter(component);
            var rewrittenCode = rewriter.GetText();
            Assert.AreEqual(expectedCode, rewrittenCode);
        }


        [TestMethod]
        [TestCategory("QuickFixes")]
        public void LabelNotUsed_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
label1:
End Sub";

            const string expectedCode =
@"Sub Foo()

End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void LabelNotUsed_QuickFixWorks_MultipleLabels()
        {
            const string inputCode =
@"Sub Foo()
label1:
dim var1 as variant
label2:
goto label1:
End Sub";

            const string expectedCode =
@"Sub Foo()
label1:
dim var1 as variant

goto label1:
End Sub"; ;

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new LineLabelNotUsedInspection(state);
            new RemoveUnusedDeclarationQuickFix(state).Fix(inspection.GetInspectionResults().First());

            var rewriter = state.GetRewriter(component);
            Assert.AreEqual(expectedCode, rewriter.GetText());
        }

    }
}
