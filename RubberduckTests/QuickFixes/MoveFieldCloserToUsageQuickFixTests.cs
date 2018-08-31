using System.Linq;
using System.Threading;
using NUnit.Framework;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.UI;
using RubberduckTests.Mocks;
using Rubberduck.Interaction;

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
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new MoveFieldCloserToUsageQuickFix(vbe.Object, state, new Mock<IMessageBox>().Object).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void MoveFieldCloserToUsage_QuickFixWorks_SingleLineIfStatemente()
        {
            const string inputCode =
                @"Private bar As String

Public Sub Foo()
    If bar = ""test"" Then Baz Else Foobar
End Sub

Private Sub Baz()
End Sub

Private Sub FooBar()
End Sub
";

            const string expectedCode =
                @"
Public Sub Foo()
    Dim bar As String
    If bar = ""test"" Then Baz Else Foobar
End Sub

Private Sub Baz()
End Sub

Private Sub FooBar()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new MoveFieldCloserToUsageQuickFix(vbe.Object, state, new Mock<IMessageBox>().Object).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void MoveFieldCloserToUsage_QuickFixWorks_SingleLineThenStatemente()
        {
            const string inputCode =
                @"Private bar As String

Public Sub Foo()
    If True Then bar = ""test""
End Sub";

            const string expectedCode =
                @"
Public Sub Foo()
    Dim bar As String
    If True Then bar = ""test""
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new MoveFieldCloserToUsageQuickFix(vbe.Object, state, new Mock<IMessageBox>().Object).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void MoveFieldCloserToUsage_QuickFixWorks_SingleLineElseStatemente()
        {
            const string inputCode =
                @"Private bar As String

Public Sub Foo()
    If True Then Else bar = ""test""
End Sub";

            const string expectedCode =
                @"
Public Sub Foo()
    Dim bar As String
    If True Then Else bar = ""test""
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                new MoveFieldCloserToUsageQuickFix(vbe.Object, state, new Mock<IMessageBox>().Object).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }
    }
}
