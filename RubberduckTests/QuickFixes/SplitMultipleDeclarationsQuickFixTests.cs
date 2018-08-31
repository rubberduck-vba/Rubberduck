using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class SplitMultipleDeclarationsQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void MultipleDeclarations_QuickFixWorks_Variables()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer, var2 As String
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Dim var1 As Integer
Dim var2 As String

End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultipleDeclarationsInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SplitMultipleDeclarationsQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void MultipleDeclarations_QuickFixWorks_Constants()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const var1 As Integer = 9, var2 As String = ""test""
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Const var1 As Integer = 9
Const var2 As String = ""test""

End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultipleDeclarationsInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SplitMultipleDeclarationsQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void MultipleDeclarations_QuickFixWorks_StaticVariables()
        {
            const string inputCode =
                @"Public Sub Foo()
    Static var1 As Integer, var2 As String
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Static var1 As Integer
Static var2 As String

End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultipleDeclarationsInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SplitMultipleDeclarationsQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
