using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteCallingConventionInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_ReturnsResult()
        {
            const string inputCode =
@"Private Declare Sub Beep CDecl Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_DoesNotReturnResult()
        {
            const string inputCode =
@"Private Declare Sub Beep Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_ReturnsMultipleResults()
        {
            const string inputCode =
@"Private Declare Sub Beep CDecl Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)
Private Declare Sub Sleep CDecl Lib ""kernel32"" (ByVal dwMilliseconds As Long)";

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_ReturnsResults_SomeObsoleteCallingConventions()
        {
            const string inputCode =
@"Private Declare Sub Beep CDecl Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)
Private Declare Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore ObsoleteCallingConvention
Private Declare Sub Beep CDecl Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ObsoleteCallingConventionInspection(null);

            Assert.AreEqual(nameof(ObsoleteCallingConventionInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteCallingConventionInspection(state);
        }
    }
}
