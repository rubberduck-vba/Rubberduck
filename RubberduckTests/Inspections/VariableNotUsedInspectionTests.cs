using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class VariableNotUsedInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_MultipleVariables_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
    Dim var2 As Date
End Sub";

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableUsed_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
    var1 = ""test""

    Goo var1
End Sub

Sub Goo(ByVal arg1 As String)
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_MultipleVariables_OneAssignedAndReferenced_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As Integer
    var1 = 8

    Dim var2 As String

    Goo var1
End Sub

Sub Goo(ByVal arg1 As Integer)
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    '@Ignore VariableNotUsed
    Dim var1 As String
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_UsedInNameStatement_DoesNotReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
    Name ""foo"" As var1
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5088        
        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_AssignedButNeverReferenced_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
    var1 = ""test""
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_AssignedMultpleTimesButNeverReferenced_ReturnsResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var2 As Long
    var2 = 4
    var2 = 7
    var2 = 8
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new VariableNotUsedInspection(state);
        }
    }
}
