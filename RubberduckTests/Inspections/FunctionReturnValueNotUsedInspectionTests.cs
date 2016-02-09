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
        public void FunctionReturnValueNotUsed_ReturnsResult_ExplicitCallWithoutAssignment()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Call Foo(""Test"")
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_CallWithoutAssignment()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Foo ""Test""
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_AddressOf()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Bar()
    Bar AddressOf Foo
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_ReturnsResult_NoReturnValueAssignment()
        {
            const string inputCode =
@"Public Function Foo() As Integer
End Function
Public Sub Bar()
    Foo
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

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_MultipleConsecutiveCalls()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Baz()
    Foo Foo(Foo(""Bar""))
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
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_IfStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Baz()
    If Foo(""Test"") Then
    End If
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
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ForEachStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    For Each Bar In Foo
    Next Bar
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
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_WhileStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    While Foo
    Wend
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
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_DoUntilStatement()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
    Do Until Foo
    Loop
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
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ReturnValueAssignment()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Public Sub Baz()
    TestVal = Foo(""Test"")
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
        public void FunctionReturnValueNotUsed_DoesNotReturnResult_ArgumentFunctionCall()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar As String) As Integer
    Foo = 42
End Function
Sub Bar(ByVal fizz As Boolean)
End Sub
Public Sub Baz()
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
