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
    public class MultilineParameterInspectionTests
    {
        private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(0, 1);

        void State_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State == ParserState.Ready)
            {
                _semaphore.Release();
            }
        }

        [TestMethod, Timeout(1000)]
        public void MultilineParameter_ReturnsResult()
        {
            const string inputCode =
@"Public Sub Foo(ByVal _
    Var1 _
    As _
    Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var inspection = new MultilineParameterInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod, Timeout(1000)]
        public void MultilineParameter_DoesNotReturnResult()
        {
            const string inputCode =
@"Public Sub Foo(ByVal Var1 As Integer)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var inspection = new MultilineParameterInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod, Timeout(1000)]
        public void MultilineParameter_ReturnsMultipleResults()
        {
            const string inputCode =
@"Public Sub Foo( _
    ByVal _
    Var1 _
    As _
    Integer, _
    ByVal _
    Var2 _
    As _
    Date)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var inspection = new MultilineParameterInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod, Timeout(1000)]
        public void MultilineParameter_ReturnsResults_SomeParams()
        {
            const string inputCode =
@"Public Sub Foo(ByVal _
    Var1 _
    As _
    Integer, ByVal Var2 As Date)
End Sub";

            //Arrange
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(inputCode, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupAllProperties();
            var parser = new RubberduckParser(vbe.Object, new RubberduckParserState());

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var inspection = new MultilineParameterInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod, Timeout(1000)]
        public void MultilineParameter_QuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo( _
    ByVal _
    Var1 _
    As _
    Integer)
End Sub";

            const string expectedCode =
@"Public Sub Foo( _
    ByVal Var1 As Integer)
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

            parser.State.StateChanged += State_StateChanged;
            parser.State.OnParseRequested();
            _semaphore.Wait();
            parser.State.StateChanged -= State_StateChanged;

            var inspection = new MultilineParameterInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod, Timeout(1000)]
        public void InspectionType()
        {
            var inspection = new MultilineParameterInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod, Timeout(1000)]
        public void InspectionName()
        {
            const string inspectionName = "MultilineParameterInspection";
            var inspection = new MultilineParameterInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}