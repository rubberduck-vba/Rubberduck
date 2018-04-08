using System.Linq;
using NUnit.Framework;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ChangeProcedureToFunctionQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Private Function Foo(ByVal arg1 As Integer) As Integer
    Foo = arg1
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks_NoAsTypeClauseInParam()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1)
End Sub";

            const string expectedCode =
                @"Private Function Foo(ByVal arg1) As Variant
    Foo = arg1
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks_DoesNotInterfereWithBody()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1 As Integer)
    arg1 = 6
    Goo arg1
End Sub

Sub Goo(ByVal a As Integer)
End Sub";

            const string expectedCode =
                @"Private Function Foo(ByVal arg1 As Integer) As Integer
    arg1 = 6
    Goo arg1
    Foo = arg1
End Function

Sub Goo(ByVal a As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks_DoesNotInterfereWithBody_BodyOnSingleLine()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1 As Integer): arg1 = 6: Goo arg1: End Sub

Sub Goo(ByVal a As Integer)
End Sub";

            const string expectedCode =
                @"Private Function Foo(ByVal arg1 As Integer) As Integer: arg1 = 6: Goo arg1:     Foo = arg1
End Function

Sub Goo(ByVal a As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks_DoesNotInterfereWithBody_BodyOnSingleLine_BodyHasStringLiteralContainingColon()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1 As Integer): arg1 = 6: Goo ""test: test"": End Sub

Sub Goo(ByVal a As String)
End Sub";

            const string expectedCode =
                @"Private Function Foo(ByVal arg1 As Integer) As Integer: arg1 = 6: Goo ""test: test"":     Foo = arg1
End Function

Sub Goo(ByVal a As String)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks_UpdatesCallsAbove()
        {
            const string inputCode =
                @"Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    Foo fizz
End Sub

Private Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    fizz = Foo(fizz)
End Sub

Private Function Foo(ByVal arg1 As Integer) As Integer
    Foo = arg1
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks_UpdatesCallsBelow()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1 As Integer)
End Sub

Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    Foo fizz
End Sub";

            const string expectedCode =
                @"Private Function Foo(ByVal arg1 As Integer) As Integer
    Foo = arg1
End Function

Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    fizz = Foo(fizz)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ChangeProcedureToFunctionQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
