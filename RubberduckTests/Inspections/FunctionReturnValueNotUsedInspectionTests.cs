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
    public class FunctionReturnValueNotUsedInspectionTests
    {
        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_NoUsages()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Function Baz() As Integer
End Function
Public Sub Bar()
    Call Foo(""Test"")
    Foo ""Test""
    Bar AddressOf Foo
    Baz
End Sub";
            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_WithUsages()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    Foo Foo(Foo(""Bar""))
    If Foo(""Test"") Then
    End If
    For Each Bar In Foo
    Next Bar
    While Foo
    Wend
    Do Until Foo
    Loop
    TestVal = Foo(""Test"")
    Bar Foo(""Test"")
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_IgnoresBuiltInFunctions()
        {
            const string inputCode =
@"Public Sub Dummy()
    MsgBox ""Test""
    Workbooks.Add
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_QuickFixWorks()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Boolean
    If True Then
        Foo = _
        True
    Else
        Foo = False
    End If
End Function";

            const string expectedCode =
@"Public Sub Foo(ByVal bar As String)
    If True Then
    Else
    End If
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

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new FunctionReturnValueNotUsedInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            string actual = module.Lines();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new FunctionReturnValueNotUsedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "FunctionReturnValueNotUsedInspection";
            var inspection = new FunctionReturnValueNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
