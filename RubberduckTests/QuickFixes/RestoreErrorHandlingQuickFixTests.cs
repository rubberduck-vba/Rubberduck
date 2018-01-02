using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RestoreErrorHandlingQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void UnhandledOnErrorResumeNext_QuickFixWorks_Procedure()
        {
            const string inputCode =
                @"Sub Foo()
    On Error Resume Next
End Sub";

            const string expectedCode =
                @"Sub Foo()
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RestoreErrorHandlingQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void UnhandledOnErrorResumeNext_QuickFixWorks_Function()
        {
            const string inputCode =
                @"Function Foo()
    On Error Resume Next
End Function";

            const string expectedCode =
                @"Function Foo()
    On Error GoTo ErrorHandler

    Exit Function
ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RestoreErrorHandlingQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void UnhandledOnErrorResumeNext_QuickFixWorks_Property()
        {
            const string inputCode =
                @"Property Get Foo() As String
    On Error Resume Next
End Property";

            const string expectedCode =
                @"Property Get Foo() As String
    On Error GoTo ErrorHandler

    Exit Property
ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Property";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RestoreErrorHandlingQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void UnhandledOnErrorResumeNext_QuickFixWorks_ExistingLabel()
        {
            const string inputCode =
                @"Sub Foo()
    On Error Resume Next

ErrorHandler:
End Sub";

            const string expectedCode =
                @"Sub Foo()
    On Error GoTo ErrorHandler1

ErrorHandler:

    Exit Sub
ErrorHandler1:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RestoreErrorHandlingQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void UnhandledOnErrorResumeNext_QuickFixWorks_MultipleIssues()
        {
            const string inputCode =
                @"Sub Foo()
    On Error Resume Next
    On Error Resume Next
End Sub";

            const string expectedCode =
                @"Sub Foo()
    On Error GoTo ErrorHandler
    On Error GoTo ErrorHandler1

    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If

    Exit Sub
ErrorHandler1:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var quickFix = new RestoreErrorHandlingQuickFix(state);

                foreach (var result in inspector.FindIssuesAsync(state, CancellationToken.None).Result)
                {
                    quickFix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void UnhandledOnErrorResumeNext_QuickFixWorks_MultipleIssuesAndExistingLabel()
        {
            const string inputCode =
                @"Sub Foo()
    On Error Resume Next
    On Error Resume Next

ErrorHandler1:
End Sub";

            const string expectedCode =
                @"Sub Foo()
    On Error GoTo ErrorHandler
    On Error GoTo ErrorHandler2

ErrorHandler1:

    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If

    Exit Sub
ErrorHandler2:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnhandledOnErrorResumeNextInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var quickFix = new RestoreErrorHandlingQuickFix(state);

                foreach (var result in inspector.FindIssuesAsync(state, CancellationToken.None).Result)
                {
                    quickFix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }
    }
}
