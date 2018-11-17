using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ChangeProcedureToFunctionQuickFixTests : QuickFixTestBase
    {

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1 As Integer)
    arg1 = 42
End Sub";

            const string expectedCode =
                @"Private Function Foo(ByVal arg1 As Integer) As Integer
    arg1 = 42
    Foo = arg1
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ProcedureCanBeWrittenAsFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks_NoAsTypeClauseInParam()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1)
    arg1 = 42
End Sub";

            const string expectedCode =
                @"Private Function Foo(ByVal arg1) As Variant
    arg1 = 42
    Foo = arg1
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ProcedureCanBeWrittenAsFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ProcedureCanBeWrittenAsFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ProcedureCanBeWrittenAsFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ProcedureCanBeWrittenAsFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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
    arg1 = 42
End Sub";

            const string expectedCode =
                @"Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    fizz = Foo(fizz)
End Sub

Private Function Foo(ByVal arg1 As Integer) As Integer
    arg1 = 42
    Foo = arg1
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ProcedureCanBeWrittenAsFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_QuickFixWorks_UpdatesCallsBelow()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1 As Integer)
    arg1 = 42
End Sub

Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    Foo fizz
End Sub";

            const string expectedCode =
                @"Private Function Foo(ByVal arg1 As Integer) As Integer
    arg1 = 42
    Foo = arg1
End Function

Sub Goo(ByVal a As Integer)
    Dim fizz As Integer
    fizz = Foo(fizz)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ProcedureCanBeWrittenAsFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ChangeProcedureToFunctionQuickFix();
        }
    }
}
