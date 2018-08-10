using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Inspections.Concrete;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteCallingConventionInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_ReturnsResult()
        {
            const string inputCode =
@"Private Declare Sub Beep CDecl Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallingConventionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_DoesNotReturnResult()
        {
            const string inputCode =
@"Private Declare Sub Beep Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallingConventionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_ReturnsMultipleResults()
        {
            const string inputCode =
@"Private Declare Sub Beep CDecl Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)
Private Declare Sub Sleep CDecl Lib ""kernel32"" (ByVal dwMilliseconds As Long)";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallingConventionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_ReturnsResults_SomeObsoleteCallingConventions()
        {
            const string inputCode =
@"Private Declare Sub Beep CDecl Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)
Private Declare Sub Sleep Lib ""kernel32"" (ByVal dwMilliseconds As Long)";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallingConventionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallingConvention_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore ObsoleteCallingConvention
Private Declare Sub Beep CDecl Lib ""kernel32"" (dwFreq As Any, dwDuration As Any)";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallingConventionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteCallingConventionInspection";
            var inspection = new ObsoleteCallingConventionInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
