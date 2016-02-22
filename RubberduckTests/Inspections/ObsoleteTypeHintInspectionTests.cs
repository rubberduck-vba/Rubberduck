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
    public class ObsoleteTypeHintInspectionTests
    {
        [TestMethod]
        public void ObsoleteCallStatement_FieldWithLongTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo&";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_FieldWithIntegerTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo%";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_FieldWithDoubleTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo#";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_FieldWithSingleTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo!";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_FieldWithDecimalTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo@";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_FieldWithStringTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo$";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_FunctionReturnsResult()
        {
            const string inputCode =
@"Public Function Foo$(ByVal bar As Boolean)
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_PropertyGetReturnsResult()
        {
            const string inputCode =
@"Public Property Get Foo$(ByVal bar As Boolean)
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_ParameterReturnsResult()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar$) As Boolean
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_VariableReturnsResult()
        {
            const string inputCode =
@"Public Function Foo() As Boolean
    Dim buzz$
    Foo = True
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_StringValueDoesNotReturnsResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim bar As String
    bar = ""Public baz$""
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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_FieldsReturnMultipleResults()
        {
            const string inputCode =
@"Public Foo$
Public Bar$";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_Field_LongTypeHint()
        {
            const string inputCode =
@"Public Foo&";

            const string expectedCode =
@"Public Foo As Long";

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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_Field_IntegerTypeHint()
        {
            const string inputCode =
@"Public Foo%";

            const string expectedCode =
@"Public Foo As Integer";

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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_Field_DoubleTypeHint()
        {
            const string inputCode =
@"Public Foo#";

            const string expectedCode =
@"Public Foo As Double";

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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_Field_SingleTypeHint()
        {
            const string inputCode =
@"Public Foo!";

            const string expectedCode =
@"Public Foo As Single";

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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_Field_DecimalTypeHint()
        {
            const string inputCode =
@"Public Foo@";

            const string expectedCode =
@"Public Foo As Decimal";

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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_Field_StringTypeHint()
        {
            const string inputCode =
@"Public Foo$";

            const string expectedCode =
@"Public Foo As String";

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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_Function_StringTypeHint()
        {
            const string inputCode =
@"Public Function Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Function";

            const string expectedCode =
@"Public Function Foo(ByVal fizz As Integer) As String
    Foo = ""test""
End Function";

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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_PropertyGet_StringTypeHint()
        {
            const string inputCode =
@"Public Property Get Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Property";

            const string expectedCode =
@"Public Property Get Foo(ByVal fizz As Integer) As String
    Foo = ""test""
End Property";

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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_Parameter_StringTypeHint()
        {
            const string inputCode =
@"Public Sub Foo(ByVal fizz$)
    Foo = ""test""
End Sub";

            const string expectedCode =
@"Public Sub Foo(ByVal fizz As String)
    Foo = ""test""
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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_Variable_StringTypeHint()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim buzz$
End Sub";

            const string expectedCode =
@"Public Sub Foo()
    Dim buzz As String
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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_FunctionReferencesAreUpdated_StringTypeHint()
        {
            const string inputCode =
@"Public Function Foo$(ByVal bar as Boolean)
    Foo$ = ""test""
End Function

Public Sub Buzz()
    Dim bat As String
    bat = Foo$()
End Sub";

            const string expectedCode =
@"Public Function Foo(ByVal bar as Boolean) As String
    Foo = ""test""
End Function

Public Sub Buzz()
    Dim bat As String
    bat = Foo()
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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void ObsoleteCallStatement_QuickFixWorks_FunctionReferencesAreUpdated_ParentIsNotChanged_StringTypeHint()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar as Boolean) As String
    Foo$ = ""test""
End Function

Public Sub Buzz()
    Dim bat As String
    bat$ = Foo$()
End Sub";

            const string expectedCode =
@"Public Function Foo(ByVal bar as Boolean) As String
    Foo = ""test""
End Function

Public Sub Buzz()
    Dim bat As String
    bat = Foo()
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

            var inspection = new ObsoleteTypeHintInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            foreach (var inspectionResult in inspectionResults)
            {
                inspectionResult.QuickFixes.First().Fix();
            }

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new ObsoleteTypeHintInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteTypeHintInspection";
            var inspection = new ObsoleteTypeHintInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}