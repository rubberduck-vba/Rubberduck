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
    public class ObsoleteLetStatementInspectionTests
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
        public void ObsoleteLetStatement_ReturnsResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
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

            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteLetStatement_ReturnsResult_MultipleLets()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
    Let var1 = var2
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

            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteLetStatement_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    var2 = var1
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

            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteLetStatement_ReturnsResult_SomeConstantsUsed()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
    var1 = var2
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

            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        public void ObsoleteLetStatement_QuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
End Sub";

            const string expectedCode =
@"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    var2 = var1
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

            var inspection = new ObsoleteLetStatementInspection(parseResult.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void InspectionType()
        {
            var inspection = new ObsoleteLetStatementInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteLetStatementInspection";
            var inspection = new ObsoleteLetStatementInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}