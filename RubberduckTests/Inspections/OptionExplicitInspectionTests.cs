using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class OptionExplicitInspectionTests : InspectionTestsBase
    {
        private const string OptionExplicitNotSpecified = @"
Sub DoSomething()
End Sub";

        [TestCase("", 0)]
        [TestCase(OptionExplicitNotSpecified, 1)]
        [TestCase("Option Explicit", 0)]
        [Category("Inspections")]
        public void OptionExplicitInspection_VariousScenario(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [TestCase(OptionExplicitNotSpecified, OptionExplicitNotSpecified, 2)]
        [TestCase(OptionExplicitNotSpecified, "Option Explicit", 1)]
        [TestCase("'@IgnoreModule", "Option Explicit", 0)]
        [TestCase("'@Ignore OptionExplicit", "Option Explicit", 0)]
        [Category("Inspections")]
        public void OptionExplicitInspection_VariousScenario_MultiComponent(string inputCode1, string inputCode2, int expectedCount)
        {
            var modules = new(string, string, ComponentType)[] 
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
            };
            Assert.AreEqual(expectedCount, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new OptionExplicitInspection(null);

            Assert.AreEqual(nameof(OptionExplicitInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new OptionExplicitInspection(state);
        }
    }
}
