using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class DefaultProjectNameInspectionTests : InspectionTestsBase
    {
        [TestCase("VBAProject", 1)]
        [TestCase("TestProject", 0)]
        [Category("Inspections")]
        public void DefaultProjectName_Names(string projectName, int expectedCount)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, string.Empty)
                .Build();
            var vbe = builder.AddProject(project).Build().Object;

            Assert.AreEqual(expectedCount, InspectionResults(vbe).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "DefaultProjectNameInspection";
            var inspection = new DefaultProjectNameInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new DefaultProjectNameInspection(state);
        }
    }
}
