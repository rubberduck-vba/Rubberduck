using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class OptionExplicitInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void NotAlreadySpecified_ReturnsResult()
        {
            const string inputCode = @"";

            //Arrange
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new OptionExplicitInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AlreadySpecified_DoesNotReturnResult()
        {
            const string inputCode = @"Option Explicit";

            //Arrange
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new OptionExplicitInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NotAlreadySpecified_ReturnsMultipleResults()
        {
            const string inputCode = @"";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();
            
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new OptionExplicitInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void PartiallySpecified_ReturnsResults()
        {
            const string inputCode1 = @"";
            const string inputCode2 = @"Option Explicit";

            //Arrange
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new OptionExplicitInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NotAlreadySpecified_QuickFixWorks()
        {
            const string inputCode = @"";
            const string expectedCode =
@"Option Explicit

";

            //Arrange
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new OptionExplicitInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();
            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void OptionExplicit_NoIgnoreQuickFix()
        {
            //Arrange
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(string.Empty, ComponentType.ClassModule, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new OptionExplicitInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.ElementAt(0).QuickFixes.Any(q => q is IgnoreOnceQuickFix));
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new OptionExplicitInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "OptionExplicitInspection";
            var inspection = new OptionExplicitInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
