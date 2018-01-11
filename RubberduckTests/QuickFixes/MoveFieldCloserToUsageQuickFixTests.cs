using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.UI;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class MoveFieldCloserToUsageQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new MoveFieldCloserToUsageQuickFix(state, new Mock<IMessageBox>().Object).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }
    }
}
