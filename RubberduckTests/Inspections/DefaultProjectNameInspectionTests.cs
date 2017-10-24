using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class DefaultProjectNameInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void DefaultProjectName_ReturnsResult()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, string.Empty)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DefaultProjectNameInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void DefaultProjectName_DoesNotReturnResult()
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, string.Empty)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DefaultProjectNameInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new DefaultProjectNameInspection(null);
            Assert.AreEqual(CodeInspectionType.MaintainabilityAndReadabilityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "DefaultProjectNameInspection";
            var inspection = new DefaultProjectNameInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
