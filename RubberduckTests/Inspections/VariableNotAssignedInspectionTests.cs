using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class VariableNotAssignedInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1 As String
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariable_ReturnsResult_MultipleVariables()
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
        public void UnassignedVariable_DoesNotReturnResult()
        {
            const string inputCode =
                @"Function Foo() As Boolean
    Dim var1 as String
    var1 = ""test""
End Function";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariable_ReturnsResult_MultipleVariables_SomeAssigned()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1 as Integer
    var1 = 8

    Dim var2 as String
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_GivenByRefAssignment_DoesNotReturnResult()
        {
            const string inputCode = @"
Sub Foo()
    Dim var1 As String
    Bar var1
End Sub

Sub Bar(ByRef value As String)
    value = ""test""
End Sub
";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }
        [Test]
        [Category("Inspections")]
        public void VariableNotAssigned_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
'@Ignore VariableNotAssigned
Dim var1 As String
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new VariableNotAssignedInspection(null);

            Assert.AreEqual(nameof(VariableNotAssignedInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new VariableNotAssignedInspection(state);
        }
    }
}
