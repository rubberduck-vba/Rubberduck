using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteGlobalInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsResult()
        {
            const string inputCode =
                @"Global var1 As Integer";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsResult_MultipleGlobals()
        {
            const string inputCode =
                @"Global var1 As Integer
Global var2 As String";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public var1 As Integer";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_ReturnsResult_SomeConstantsUsed()
        {
            const string inputCode =
                @"Public var1 As Integer
Global var2 As Date";
        Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteGlobal_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ObsoleteGlobal
Global var1 As Integer";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ObsoleteGlobalInspection(null);

            Assert.AreEqual(nameof(ObsoleteGlobalInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteGlobalInspection(state);
        }
    }
}
