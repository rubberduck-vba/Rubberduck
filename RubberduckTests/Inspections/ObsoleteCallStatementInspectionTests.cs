using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObsoleteCallStatementInspectionTests
    {
        [TestMethod]
        public void ObsoleteCallStatement_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Call Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteCallStatementInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_DoesNotReturnResult()
        {
            const string inputCode =
@"Sub Foo()
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteCallStatementInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_ReturnsMultipleResults()
        {
            const string inputCode =
@"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteCallStatementInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_ReturnsResults_SomeObsoleteCallStatements()
        {
            const string inputCode =
@"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteCallStatementInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_RemoveCallStatement()
        {
            const string inputCode =
@"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";

            const string expectedCode =
@"Sub Foo()
    Goo 1, ""test""
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Foo
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.ParseSynchronous();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteCallStatementInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            var actual = module.Lines();
            Assert.AreEqual(expectedCode, actual);
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new ObsoleteCallStatementInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteCallStatementInspection";
            var inspection = new ObsoleteCallStatementInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}