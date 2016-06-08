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
    public class ImplicitVariantReturnTypeInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_ReturnsResult_Function()
        {
            const string inputCode =
@"Function Foo()
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_ReturnsResult_LibraryFunction()
        {
            const string inputCode =
@"Declare PtrSafe Function CreateProcess Lib ""kernel32"" _
                                   Alias ""CreateProcessA""(ByVal lpApplicationName As String, _
                                                           ByVal lpCommandLine As String, _
                                                           lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                                           lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                                           ByVal bInheritHandles As Long, _
                                                           ByVal dwCreationFlags As Long, _
                                                           lpEnvironment As Any, _
                                                           ByVal lpCurrentDirectory As String, _
                                                           lpStartupInfo As STARTUPINFO, _
                                                           lpProcessInformation As PROCESS_INFORMATION)";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_DoesNotReturnResult_LibraryFunction()
        {
            const string inputCode =
@"Declare PtrSafe Function CreateProcess Lib ""kernel32"" _
                                   Alias ""CreateProcessA""(ByVal lpApplicationName As String, _
                                                           ByVal lpCommandLine As String, _
                                                           lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                                           lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                                           ByVal bInheritHandles As Long, _
                                                           ByVal dwCreationFlags As Long, _
                                                           lpEnvironment As Any, _
                                                           ByVal lpCurrentDirectory As String, _
                                                           lpStartupInfo As STARTUPINFO, _
                                                           lpProcessInformation As PROCESS_INFORMATION) As Long";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_ReturnsResult_PropertyGet()
        {
            const string inputCode =
@"Property Get Foo()
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_ReturnsResult_MultipleFunctions()
        {
            const string inputCode =
@"Function Foo()
End Function

Function Goo()
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_DoesNotReturnResult()
        {
            const string inputCode =
@"Function Foo() As Boolean
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_ReturnsResult_MultipleSubs_SomeReturning()
        {
            const string inputCode =
@"Function Foo()
End Function

Function Goo() As String
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_QuickFixWorks_Function()
        {
            const string inputCode =
@"Function Foo()
End Function";

            const string expectedCode =
@"Function Foo() As Variant
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_QuickFixWorks_PropertyGet()
        {
            const string inputCode =
@"Property Get Foo()
End Property";

            const string expectedCode =
@"Property Get Foo() As Variant
End Property";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_QuickFixWorks_LibraryFunction()
        {
            const string inputCode =
@"Declare PtrSafe Function CreateProcess Lib ""kernel32"" _
                                   Alias ""CreateProcessA""(ByVal lpApplicationName As String, _
                                                           ByVal lpCommandLine As String, _
                                                           lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                                           lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                                           ByVal bInheritHandles As Long, _
                                                           ByVal dwCreationFlags As Long, _
                                                           lpEnvironment As Any, _
                                                           ByVal lpCurrentDirectory As String, _
                                                           lpStartupInfo As STARTUPINFO, _
                                                           lpProcessInformation As PROCESS_INFORMATION)";

            const string expectedCode =
@"Declare PtrSafe Function CreateProcess Lib ""kernel32"" _
                                   Alias ""CreateProcessA""(ByVal lpApplicationName As String, _
                                                           ByVal lpCommandLine As String, _
                                                           lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                                           lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                                           ByVal bInheritHandles As Long, _
                                                           ByVal dwCreationFlags As Long, _
                                                           lpEnvironment As Any, _
                                                           ByVal lpCurrentDirectory As String, _
                                                           lpStartupInfo As STARTUPINFO, _
                                                           lpProcessInformation As PROCESS_INFORMATION) As Variant";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitVariantReturnType_QuickFixWorks_Function_HasComment()
        {
            const string inputCode =
@"Function Foo()    ' comment
End Function";

            const string expectedCode =
@"Function Foo() As Variant    ' comment
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ImplicitVariantReturnTypeInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ImplicitVariantReturnTypeInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitVariantReturnTypeInspection";
            var inspection = new ImplicitVariantReturnTypeInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
