using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ReplaceIfElseWithConditionalStatementQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void Simple()
        {
            const string inputCode =
                @"Sub Foo()
    Dim d As Boolean
    If True Then
        d = True
    Else
        d = False
    EndIf
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    d = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ReplaceIfElseWithConditionalStatementQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ComplexCondition()
        {
            const string inputCode =
                @"Sub Foo()
    Dim d As Boolean
    If True Or False And False Xor True Then
        d = True
    Else
        d = False
    EndIf
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    d = True Or False And False Xor True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ReplaceIfElseWithConditionalStatementQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void InvertedCondition()
        {
            const string inputCode =
                @"Sub Foo()
    Dim d As Boolean
    If True Then
        d = False
    Else
        d = True
    EndIf
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    d = Not (True)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ReplaceIfElseWithConditionalStatementQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void QualifiedName()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Fizz.Buzz = True
    Else
        Fizz.Buzz = False
    EndIf
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Fizz.Buzz = True
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new ReplaceIfElseWithConditionalStatementQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }
    }
}
