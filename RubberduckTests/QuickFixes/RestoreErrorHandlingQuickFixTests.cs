using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RestoreErrorHandlingQuickFixTests : QuickFixTestBase
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UnhandledOnErrorResumeNextInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UnhandledOnErrorResumeNextInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UnhandledOnErrorResumeNextInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new UnhandledOnErrorResumeNextInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new UnhandledOnErrorResumeNextInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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
    On Error GoTo ErrorHandler2
    On Error GoTo ErrorHandler3

ErrorHandler1:

    Exit Sub
ErrorHandler2:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If

    Exit Sub
ErrorHandler3:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new UnhandledOnErrorResumeNextInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnhandledOnErrorResumeNext_QuickFixWorks_SameLabelInMultipleProcedures()
        {
            const string inputCode =
 @"Sub Foo()
    On Error Resume Next
End Sub

Sub Bar()
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
End Sub

Sub Bar()
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new UnhandledOnErrorResumeNextInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnhandledOnErrorResumeNext_QuickFixWorks_LabelWithNonNumericSuffix()
        {
            const string inputCode =
                @"Sub Foo()
    On Error Resume Next

ErrorHandlerFoo:
End Sub";

            const string expectedCode =
                @"Sub Foo()
    On Error GoTo ErrorHandler

ErrorHandlerFoo:

    Exit Sub
ErrorHandler:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new UnhandledOnErrorResumeNextInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnhandledOnErrorResumeNext_QuickFixWorks_GeneratedLabelHasGreaterNumberThanExistingLabel()
        {
            const string inputCode =
                @"Sub Foo()
    On Error Resume Next

ErrorHandler3:
End Sub";

            const string expectedCode =
                @"Sub Foo()
    On Error GoTo ErrorHandler4

ErrorHandler3:

    Exit Sub
ErrorHandler4:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new UnhandledOnErrorResumeNextInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RestoreErrorHandlingQuickFix();
        }
    }
}
