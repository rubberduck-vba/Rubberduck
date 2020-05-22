using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class OptionBaseInspectionTests : InspectionTestsBase
    {
        [TestCase("Option Base 1", 1)]
        [TestCase("", 0)]
        [TestCase("Option Base 0", 0)]
        [TestCase("'@Ignore OptionBase\r\nOption Base 1", 0)]
        [Category("Inspections")]
        public void OptionBaseInspection_VariousScenarios_StdModule(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [TestCase("Option Base 1", "Option Base 1", 2)]
        [TestCase("", "Option Base 1", 1)]
        [Category("Inspections")]
        public void OptionBaseInspection_VariousScenarios_Classes(string inputCode1, string inputCode2, int expectedCount)
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
            var inspection = new OptionBaseInspection(null);

            Assert.AreEqual(nameof(OptionBaseInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new OptionBaseInspection(state);
        }
    }
}
