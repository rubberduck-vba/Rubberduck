using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class HostSpecificExpressionInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ReturnsResultForExpressionOnLeftHandSide()
        {
            const string code = @"
Public Sub DoSomething()
    [A1] = 42
End Sub
";
            var builder = new MockVbeBuilder();
            IVBComponent component;
            var vbe = builder.BuildFromSingleStandardModule(code, out component);
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupGet(m => m.ApplicationName).Returns("Excel");
            vbe.Setup(m => m.HostApplication()).Returns(() => mockHost.Object);

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(new Mock<ISinks>().Object));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new HostSpecificExpressionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }
    }
}