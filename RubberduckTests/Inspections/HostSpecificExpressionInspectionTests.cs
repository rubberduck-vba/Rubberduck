using System.Linq;
using System.Threading;
using NUnit.Framework;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Common;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class HostSpecificExpressionInspectionTests
    {
        [Test]
        [Category("Inspections")]
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

            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                state.AddTestLibrary("VBA.4.2.xml");
                state.AddTestLibrary("Excel.1.8.xml");
                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new HostSpecificExpressionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }
    }
}