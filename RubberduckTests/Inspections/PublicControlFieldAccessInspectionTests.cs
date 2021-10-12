using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class PublicControlFieldAccessInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ReferencedInsideContainingForm_NoResult()
        {
            var formCode = @"
Option Explicit
Public Sub Test()
    Me.TextBox1.Text = ""TEST""
End Sub";
            var moduleCode = @"
Option Explicit";

            var mockVbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, moduleCode)
                .MockUserFormBuilder("UserForm1", formCode)
                .AddControl("TextBox1")
                .AddFormToProjectBuilder()
                .AddProjectToVbeBuilder()
                .Build();
            var results = InspectionResults(mockVbe.Object);

            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ReferencedOutsideContainingForm_HasResult()
        {
            var formCode = @"
Option Explicit";
            var moduleCode = @"
Option Explicit

Public Sub Test()
    With New UserForm1
        MsgBox .TextBox1.Text
    End With
End Sub
";

            var mockVbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, moduleCode)
                .MockUserFormBuilder("UserForm1", formCode)
                .AddControl("TextBox1")
                .AddFormToProjectBuilder()
                .AddProjectToVbeBuilder()
                .Build();
            var results = InspectionResults(mockVbe.Object);

            Assert.AreEqual(1, results.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new PublicControlFieldAccessInspection(state);
        }
    }
}