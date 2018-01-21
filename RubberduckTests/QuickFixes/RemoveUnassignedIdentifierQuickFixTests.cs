using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnassignedIdentifierQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 as Integer
End Sub";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotAssignedInspection(state);
                new RemoveUnassignedIdentifierQuickFix(state).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_VariableOnMultipleLines_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim _
var1 _
as _
Integer
End Sub";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotAssignedInspection(state);
                new RemoveUnassignedIdentifierQuickFix(state).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_MultipleVariablesOnSingleLine_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As Integer, var2 As Boolean
End Sub";

            const string expectedCode =
                @"Sub Foo()
Dim var1 As Integer
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotAssignedInspection(state);
                new RemoveUnassignedIdentifierQuickFix(state).Fix(
                    inspection.GetInspectionResults().Single(s => s.Target.IdentifierName == "var2"));

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_MultipleVariablesOnMultipleLines_QuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As Integer, _
var2 As Boolean
End Sub";

            const string expectedCode =
                @"Sub Foo()
Dim var1 As Integer
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotAssignedInspection(state);
                new RemoveUnassignedIdentifierQuickFix(state).Fix(
                    inspection.GetInspectionResults().Single(s => s.Target.IdentifierName == "var2"));

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }
    }
}
