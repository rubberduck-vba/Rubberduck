using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObsoleteCommentSyntaxInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_ReturnsResult()
        {
            const string inputCode = @"Rem test";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_DoesNotReturnResult_QuoteComment()
        {
            const string inputCode = @"' test";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_DoesNotReturnResult_OtherParseInspectionFires()
        {
            const string inputCode = @"
Sub foo()
    Dim i As String
    i = """"
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var emptyStringLiteralInspection = new EmptyStringLiteralInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection, emptyStringLiteralInspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count(r => r.Inspection is ObsoleteCommentSyntaxInspection));
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_DoesNotReturnResult_RemInStringLiteral()
        {
            const string inputCode =
                @"Sub Foo()
    Dim bar As String
    bar = ""iejo rem oernp"" ' test
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_ReturnsMultipleResults()
        {
            const string inputCode =
                @"Rem test1
Rem test2";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_ReturnsResults_SomeObsoleteCommentSyntax()
        {
            const string inputCode =
                @"Rem test1
' test2";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_Ignored_DoesNotReturnResult()
        {
            const string inputCode = @"
'@Ignore ObsoleteCommentSyntax
Rem test";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ObsoleteCommentSyntaxInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteCommentSyntaxInspection";
            var inspection = new ObsoleteCommentSyntaxInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
