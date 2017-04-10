using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using ParserState = Rubberduck.Parsing.VBA.ParserState;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class MemberNotOnInterfaceInspectionTests
    {
        private static ParseCoordinator ArrangeParser(string inputCode, string library = "Scripting")
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Codez", ComponentType.StandardModule, inputCode)
                .AddReference(library, 
                              library.Equals("Scripting") ? MockVbeBuilder.LibraryPathScripting : MockVbeBuilder.LibraryPathMsExcel,
                              1,
                              library.Equals("Scripting") ? 0 : 8,
                              true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));

            parser.State.AddTestLibrary(library.Equals("Scripting") ? "Scripting.1.0.xml" : "Excel.1.8.xml");
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
        public void MemberNotOnInterface_ReturnsResult_ApplicationObject()
        {
            const string inputCode =
@"Sub Foo()
    Application.NonMember
End Sub";

            var parser = ArrangeParser(inputCode, "Excel");

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
            const string inputCode =
@"Sub Foo()
    Dim dict As New Dictionary
    With dict
        .NonMember
    End With
End Sub";

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

            var parser = ArrangeParser(inputCode);

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new MemberNotOnInterfaceInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }
    }
}
