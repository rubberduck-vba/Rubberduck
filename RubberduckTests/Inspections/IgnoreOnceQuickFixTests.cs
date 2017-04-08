using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class IgnoreOnceQuickFixTests
    {
        [TestMethod]
        public void AnnotationListFollowedByCommentAddsAnnotationCorrectly()
        {
            const string inputCode = @"
Public Function GetSomething() As Long
    '@Ignore VariableNotAssigned: Is followed by a comment.
    Dim foo
    GetSomething = foo
End Function
";

            const string expectedCode = @"
Public Function GetSomething() As Long
    '@Ignore VariableTypeNotDeclared, VariableNotAssigned: Is followed by a comment.
    Dim foo
    GetSomething = foo
End Function
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableTypeNotDeclaredInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

    }
}
