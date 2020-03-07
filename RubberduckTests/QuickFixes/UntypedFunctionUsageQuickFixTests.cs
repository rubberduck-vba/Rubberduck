using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;


namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class UntypedFunctionUsageQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        [Ignore("Broken feature - passes locally but not in AV. See FIXME in the notes")]
        public void UntypedFunctionUsage_QuickFixWorks()
        {
            const string inputCode =
@"Sub Foo()
    Dim str As String
    str = Left(""test"", 1)
End Sub";

            const string expectedCode =
@"Sub Foo()
    Dim str As String
    str = Left$(""test"", 1)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .AddReference(ReferenceLibrary.VBA)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var component = project.Object.VBComponents[0];
            var (parser, rewriteManager) = MockParser.CreateWithRewriteManager(vbe.Object);
            using (var state = parser.State)
            {
                // FIXME reinstate and unignore tests
                // refers to "UntypedFunctionUsageInspectionTests.GetBuiltInDeclarations()"
                //GetBuiltInDeclarations().ForEach(d => parser.State.AddDeclaration(d));

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new UntypedFunctionUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);

                var rewriteSession = rewriteManager.CheckOutCodePaneSession();

                new UntypedFunctionUsageQuickFix().Fix(inspectionResults.First(), rewriteSession);

                var actualCode = rewriteSession.CheckOutModuleRewriter(component.QualifiedModuleName).GetText();
                Assert.AreEqual(expectedCode, actualCode);
            }
        }
    }
}
