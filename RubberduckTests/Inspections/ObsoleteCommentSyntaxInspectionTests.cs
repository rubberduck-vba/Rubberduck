using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
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
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_DoesNotReturnResult_QuoteComment()
        {
            const string inputCode = @"' test";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
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
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var emptyStringLiteralInspection = new EmptyStringLiteralInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection, emptyStringLiteralInspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count(r => r.Inspection is ObsoleteCommentSyntaxInspection));
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
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_ReturnsMultipleResults()
        {
            const string inputCode =
@"Rem test1
Rem test2";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_ReturnsResults_SomeObsoleteCommentSyntax()
        {
            const string inputCode =
@"Rem test1
' test2";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_Ignored_DoesNotReturnResult()
        {
            const string inputCode = @"
'@Ignore ObsoleteCommentSyntax
Rem test";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment()
        {
            const string inputCode =
@"Rem test1";

            const string expectedCode =
@"' test1";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ReplaceObsoleteCommentMarkerQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateCommentHasContinuation()
        {
            const string inputCode =
@"Rem this is _
a comment";

            const string expectedCode =
@"' this is _
a comment";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ReplaceObsoleteCommentMarkerQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment()
        {
            const string inputCode =
@"Rem test1";

            const string expectedCode =
@"";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveCommentQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveCommentHasContinuation()
        {
            const string inputCode =
@"Rem test1 _
continued";

            const string expectedCode =
@"";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveCommentQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment_LineHasCode()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is a comment";

            const string expectedCode =
@"Dim Foo As Integer: ' This is a comment";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ReplaceObsoleteCommentMarkerQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_UpdateComment_LineHasCodeAndContinuation()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is _
a comment";

            const string expectedCode =
@"Dim Foo As Integer: ' This is _
a comment";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new ReplaceObsoleteCommentMarkerQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment_LineHasCode()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is a comment";

            const string expectedCode =
@"Dim Foo As Integer: ";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveCommentQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_QuickFixWorks_RemoveComment_LineHasCodeAndContinuation()
        {
            const string inputCode =
@"Dim Foo As Integer: Rem This is _
a comment";

            const string expectedCode =
@"Dim Foo As Integer: ";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveCommentQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteCommentSyntax_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Rem test1";

            const string expectedCode =
@"'@Ignore ObsoleteCommentSyntax
Rem test1";
            
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteCommentSyntaxInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
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
