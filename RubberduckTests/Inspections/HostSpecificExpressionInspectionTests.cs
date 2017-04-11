using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class HostSpecificExpressionInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        [DeploymentItem(@"TestFiles\")]
        public void ReturnsResultForExpressionOnLeftHandSide()
        {
            const string code = @"
Public Sub DoSomething()
    [A1] = 42
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, code)
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 2, true)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 7, true)
                .Build();
            var vbe = builder.AddProject(project).Build();
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupGet(m => m.ApplicationName).Returns("Excel");
            vbe.Setup(m => m.HostApplication()).Returns(() => mockHost.Object);

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.State.AddTestLibrary("VBA.4.2.xml");
            parser.State.AddTestLibrary("Excel.1.8.xml");
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new HostSpecificExpressionInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }
    }
}