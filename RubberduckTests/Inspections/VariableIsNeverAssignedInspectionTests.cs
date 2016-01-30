using System.Linq;
using System.Threading;
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
    public class VariableIsNeverAssignedInspectionTests
    {
        private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(0, 1);

        void State_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State == ParserState.Ready)
            {
                _semaphore.Release();
            }
        }

        [TestMethod]
        public void VariableNotAssigned_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new VariableNotAssignedInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void UnassignedVariable_ReturnsResult_MultipleVariables()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 As String
    Dim var2 As Date
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new VariableNotAssignedInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void UnassignedVariable_DoesNotReturnResult()
        {
            const string inputCode =
@"Function Foo() As Boolean
    Dim var1 as String
    var1 = ""test""
End Function";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new VariableNotAssignedInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void UnassignedVariable_ReturnsResult_MultipleVariables_SomeAssigned()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 as Integer
    var1 = 8

    Dim var2 as String
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new VariableNotAssignedInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void UnassignedVariable_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
    Dim var1 as Integer
End Sub";

            const string expectedCode =
@"Sub Foo()
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var project = vbe.Object.VBProjects.Item(0);
            var module = project.VBComponents.Item(0).CodeModule;
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parseResult = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parseResult.State.StateChanged += State_StateChanged;
            parseResult.State.OnParseRequested();
            _semaphore.Wait();
            parseResult.State.StateChanged -= State_StateChanged;

            var inspection = new VariableNotAssignedInspection(parseResult.State);
            inspection.GetInspectionResults().First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new VariableNotAssignedInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "VariableNotAssignedInspection";
            var inspection = new VariableNotAssignedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}