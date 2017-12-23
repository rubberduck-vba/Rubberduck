using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class PassParameterByReferenceQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_PassByReferenceQuickFixWorks()
        {

            string inputCode =
@"Public Sub Foo(Optional ByVal barByVal As String = ""XYZ"")
    Let barByVal = ""test""
End Sub";
            string expectedCode =
@"Public Sub Foo(Optional ByRef barByVal As String = ""XYZ"")
    Let barByVal = ""test""
End Sub";

            var quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //check when ByVal argument is one of several parameters
            inputCode =
@"Public Sub Foo(ByRef firstArg As Long, Optional ByVal barByVal As String = """", secondArg as Double)
    Let barByVal = ""test""
End Sub";
            expectedCode =
@"Public Sub Foo(ByRef firstArg As Long, Optional ByRef barByVal As String = """", secondArg as Double)
    Let barByVal = ""test""
End Sub";

            quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
            //AppleWatch IDE test
            inputCode =
@"
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
"
;
            expectedCode =
@"
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
"
;
            quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            inputCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal _xByValbar As Long,  ByVal _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            expectedCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal _xByValbar As Long,  ByRef _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";

            quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            inputCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByVal _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            expectedCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByRef _
    barTwo _
    As _
    Long)
barTwo = 42
End Sub
";

            quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            inputCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByVal barTwo _
    As _
    Long)
barTwo = 42
End Sub
";
            expectedCode =
@"Private Sub Foo(ByVal barByVal As Long, ByVal barTwoon As Long,  ByRef barTwo _
    As _
    Long)
barTwo = 42
End Sub
";

            quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
            //weaponized code test
            inputCode =
@"Sub DoSomething(_
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

            expectedCode =
@"Sub DoSomething(_
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
            quickFixResult = ApplyPassParameterByReferenceQuickFixToVBAFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        private string ApplyPassParameterByReferenceQuickFixToVBAFragment(string inputCode)
        {
            var vbe = BuildMockVBEStandardModuleForVBAFragment(inputCode);
            using(var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new AssignedByValParameterInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new PassParameterByReferenceQuickFix(state).Fix(inspectionResults.First());
                return state.GetRewriter(vbe.Object.ActiveVBProject.VBComponents[0]).GetText();
            }
        }

        private Mock<IVBE> BuildMockVBEStandardModuleForVBAFragment(string inputCode)
        {
            return MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
        }
    }
}
