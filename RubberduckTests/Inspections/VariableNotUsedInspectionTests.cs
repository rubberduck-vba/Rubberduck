using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class VariableNotUsedInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
End Sub";
            Assert.AreEqual(1, GetTestResultCount(inputCode));
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

            Assert.AreEqual(2, GetTestResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void VariableUsed_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 as String
    var1 = ""test""

    Goo var1
End Sub

Sub Goo(ByVal arg1 As String)
End Sub";

            Assert.AreEqual(0, GetTestResultCount(inputCode));
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_MultipleVariables_OneAssignedAndReferenced_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 as Integer
    var1 = 8

    Dim var2 as String

    Goo var1
End Sub

Sub Goo(ByVal arg1 As Integer)
End Sub";

            Assert.AreEqual(1, GetTestResultCount(inputCode));
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

            Assert.AreEqual(0, GetTestResultCount(inputCode));
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

            Assert.AreEqual(0, GetTestResultCount(inputCode));
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

            Assert.AreEqual(1, GetTestResultCount(inputCode));
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

            Assert.AreEqual(1, GetTestResultCount(inputCode));
        }

        private int GetTestResultCount(string inputCode)
        {
            var resultCount = 0;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                resultCount = inspectionResults.Count();
            }
            return resultCount;
        }
    }
}
