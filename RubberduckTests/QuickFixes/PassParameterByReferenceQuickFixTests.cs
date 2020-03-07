using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class PassParameterByReferenceQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_PassByReferenceQuickFixWorks()
        {

            const string inputCode = @"Public Sub Foo(Optional ByVal barByVal As String = ""XYZ"")
    Let barByVal = ""test""
End Sub";
            const string expectedCode = @"Public Sub Foo(Optional ByRef barByVal As String = ""XYZ"")
    Let barByVal = ""test""
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AssignedByValParameterInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_PassByReferenceQuickFixWorks_ByValParameterIsOneOfSeveral()
        {
            const string inputCode = @"Public Sub Foo(ByRef firstArg As Long, Optional ByVal barByVal As String = """", secondArg as Double)
    Let barByVal = ""test""
End Sub";
            const string expectedCode = @"Public Sub Foo(ByRef firstArg As Long, Optional ByRef barByVal As String = """", secondArg as Double)
    Let barByVal = ""test""
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AssignedByValParameterInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_PassByReferenceQuickFixWorks_LineContinued1()
        {
            const string inputCode = @"
Private Sub Foo(Optional ByVal  _
    bar _
    As _
    Long = 4, _
    ByVal _
    barTwo _
    As _
    Long)
bar = 42
End Sub
";
            const string expectedCode = @"
Private Sub Foo(Optional ByRef  _
    bar _
    As _
    Long = 4, _
    ByVal _
    barTwo _
    As _
    Long)
bar = 42
End Sub
";
            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AssignedByValParameterInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_PassByReferenceQuickFixWorks_LineContinued2()
        {
            const string inputCode = @"Private Sub Foo(ByVal barByVal As Long, ByVal _xByValbar As Long,  ByVal _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            const string expectedCode = @"Private Sub Foo(ByVal barByVal As Long, ByVal _xByValbar As Long,  ByRef _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AssignedByValParameterInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_PassByReferenceQuickFixWorks_LineContinued3()
        {
            const string inputCode = @"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByVal _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            const string expectedCode = @"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByRef _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AssignedByValParameterInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_PassByReferenceQuickFixWorks_LineContinued4()
        {
            const string inputCode = @"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByVal barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            const string expectedCode = @"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByRef barTwo _
    As _
    Long)
barTwo = 42
End Sub
";


            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AssignedByValParameterInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_PassByReferenceQuickFixWorks_LineContinued5()
        {
            //weaponized code test
            const string inputCode = @"Sub DoSomething( _
    ByVal foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
    foo = 4
    bar = barbecue * _
               bar + foo / barbecue
End Sub
";

            const string expectedCode = @"Sub DoSomething( _
    ByRef foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
    foo = 4
    bar = barbecue * _
               bar + foo / barbecue
End Sub
";
            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new AssignedByValParameterInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new PassParameterByReferenceQuickFix();
        }
    }
}
