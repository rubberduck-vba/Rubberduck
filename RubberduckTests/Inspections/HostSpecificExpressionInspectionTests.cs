using System.Linq;
using NUnit.Framework;
using Moq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class HostSpecificExpressionInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ReturnsResultForExpressionOnLeftHandSide()
        {
            const string code = @"
Public Sub DoSomething()
    [A1] = 42
End Sub
";
            var vbe = MockVbeBuilder.BuildFromModules(("Module1", code, ComponentType.StandardModule), new ReferenceLibrary[] { ReferenceLibrary.VBA, ReferenceLibrary.Excel });
            var mockHost = new Mock<IHostApplication>();
            mockHost.SetupGet(m => m.ApplicationName).Returns("Excel");
            vbe.Setup(m => m.HostApplication()).Returns(() => mockHost.Object);

            Assert.AreEqual(1, InspectionResults(vbe.Object).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new HostSpecificExpressionInspection(null);

            Assert.AreEqual(nameof(HostSpecificExpressionInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new HostSpecificExpressionInspection(state);
        }
    }
}