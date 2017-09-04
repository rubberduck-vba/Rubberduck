using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class MoveFieldCloserToUsageQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        public void MoveFieldCloserToUsage_QuickFixWorks()
        {
            const string inputCode =
@"Private bar As String
Public Sub Foo()
    bar = ""test""
End Sub";

            const string expectedCode =
@"Public Sub Foo()
    Dim bar As String
bar = ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new MoveFieldCloserToUsageInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new MoveFieldCloserToUsageQuickFix(state, new Mock<IMessageBox>().Object).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }
    }
}
