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
    Dim foo As Long
    GetSomething = foo
End Function
";

            const string expectedCode = @"
Public Function GetSomething() As Long
    '@Ignore UnassignedVariableUsage, VariableNotAssigned: Is followed by a comment.
    Dim foo As Long
    GetSomething = foo
End Function
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new UnassignedVariableUsageInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

    }
}
