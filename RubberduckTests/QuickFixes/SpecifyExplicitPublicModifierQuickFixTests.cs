using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;


namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class SpecifyExplicitPublicModifierQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ImplicitPublicMember_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 as Integer)
'Just an inoffensive little comment

End Sub";

            const string expectedCode =
@"Public Sub Foo(ByVal arg1 as Integer)
'Just an inoffensive little comment

End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitPublicMemberInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new SpecifyExplicitPublicModifierQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

    }
}
