using System;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnassignedIdentifierQuickFixTests : QuickFixTestBase
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotAssignedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableNotAssignedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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
            Func<IInspectionResult, bool> conditionToFix = s => s.Target.IdentifierName == "var2";
            var actualCode = ApplyQuickFixToFirstInspectionResultSatisfyingPredicate(inputCode, state => new VariableNotAssignedInspection(state), conditionToFix);
            Assert.AreEqual(expectedCode, actualCode);
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

            Func<IInspectionResult, bool> conditionToFix = result => result.Target.IdentifierName == "var2";
            var actualCode = ApplyQuickFixToFirstInspectionResultSatisfyingPredicate(inputCode, state => new VariableNotAssignedInspection(state), conditionToFix);
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveUnassignedIdentifierQuickFix();
        }
    }
}
