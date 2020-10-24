using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class QualifyWithMeQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void QualifiesImplicitWorkbookReferencesInWorkbooks()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Worksheets(""Sheet1"")
End Sub";

            const string expectedCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Me.Worksheets(""Sheet1"")
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResultForImplicitWorkbookInspection(inputCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        private string ApplyQuickFixToFirstInspectionResultForImplicitWorkbookInspection(string inputCode)
        {
            var inputModule = ("SomeWorkbook", inputCode, ComponentType.Document);
            var vbe = MockVbeBuilder.BuildFromModules(inputModule, ReferenceLibrary.Excel).Object;

            var (state, rewriteManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var documentModule = state.DeclarationFinder.UserDeclarations(DeclarationType.Document)
                    .OfType<DocumentModuleDeclaration>()
                    .Single();
                documentModule.AddSupertypeName("Workbook");

                var inspection = new ImplicitContainingWorkbookReferenceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                var rewriteSession = rewriteManager.CheckOutCodePaneSession();

                var quickFix = QuickFix(state);

                var resultToFix = inspectionResults.First();
                quickFix.Fix(resultToFix, rewriteSession);

                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName == "SomeWorkbook");

                return rewriteSession.CheckOutModuleRewriter(module).GetText();
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void QualifiesImplicitWorksheetReferencesInWorksheets()
        {
            const string inputCode =
                @"
Private Sub Example()
    Dim foo As Range
    Set foo = Range(""A1"") 
End Sub";

            const string expectedCode =
                @"
Private Sub Example()
    Dim foo As Range
    Set foo = Me.Range(""A1"") 
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResultForImplicitWorksheetInspection(inputCode);
            Assert.AreEqual(expectedCode, actualCode);
        }

        private string ApplyQuickFixToFirstInspectionResultForImplicitWorksheetInspection(string inputCode)
        {
            var inputModule = ("Sheet1", inputCode, ComponentType.Document);
            var vbe = MockVbeBuilder.BuildFromModules(inputModule, ReferenceLibrary.Excel).Object;

            var (state, rewriteManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var documentModule = state.DeclarationFinder.UserDeclarations(DeclarationType.Document)
                    .OfType<DocumentModuleDeclaration>()
                    .Single();
                documentModule.AddSupertypeName("Worksheet");

                var inspection = new ImplicitContainingWorksheetReferenceInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                var rewriteSession = rewriteManager.CheckOutCodePaneSession();

                var quickFix = QuickFix(state);

                var resultToFix = inspectionResults.First();
                quickFix.Fix(resultToFix, rewriteSession);

                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName == "Sheet1");

                return rewriteSession.CheckOutModuleRewriter(module).GetText();
            }
        }

        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new QualifyWithMeQuickFix();
        }
    }
}
