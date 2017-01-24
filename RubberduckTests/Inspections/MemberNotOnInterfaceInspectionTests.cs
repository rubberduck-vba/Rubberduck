using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using ParserState = Rubberduck.Parsing.VBA.ParserState;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class MemberNotOnInterfaceInspectionTests
    {
        private static ParseCoordinator ArrangeParser(string inputCode)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Codez", ComponentType.StandardModule, inputCode)
                .AddReference("Scripting", MockVbeBuilder.LibraryPathScripting, 1, 0, true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));

            parser.State.AddTestLibrary("Scripting.1.0.xml");
            return parser;
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_UnDeclaredMember()
        {
            const string inputCode =
@"Sub Foo()
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict.NonMember
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_UnDeclaredInterfaceMember()
        {
            const string inputCode =
@"Sub Foo()
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict.NonMember
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_UnDeclaredMemberOnParameter()
        {
            const string inputCode =
@"Sub Foo(dict As Dictionary)
    dict.NonMember
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_DeclaredMember()
        {            
            const string inputCode =
@"Sub Foo()
    Dim dict As Dictionary
    Set dict = New Dictionary
    Debug.Print dict.Count
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_NonExtensible()
        {
            const string inputCode =
@"Sub Foo()
    Dim x As File
    Debug.Print x.NonMember
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_ReturnsResult_WithBlock()
        {
            Assert.Inconclusive("This is currently not working.");
            const string inputCode =
@"Sub Foo()
    Dim dict As New Dictionary
    With dict
        .NonMember
    End With
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_BangNotation()
        {
            const string inputCode =
@"Sub Foo()
    Dim dict As Dictionary
    Set dict = New Dictionary
    dict!SomeIdentifier = 42
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_WithBlockBangNotation()
        {
            const string inputCode =
@"Sub Foo()
    Dim dict As New Dictionary
    With dict
        !SomeIdentifier = 42
    End With
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_DoesNotReturnResult_ProjectReference()
        {
            const string inputCode =
@"Sub Foo()
    Dim dict As Scripting.Dictionary
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void MemberNotOnInterface_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo(dict As Dictionary)
    Dim dict As Dictionary
    Set dict = New Dictionary
    '@Ignore MemberNotOnInterface
    dict.NonMember
End Sub";

            //Arrange
            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }
    }
}
