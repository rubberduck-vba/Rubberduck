using System.Linq;
using System.Threading;
using NUnit.Framework;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;
using Rubberduck.Interaction;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveUnusedParameterQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        [Apartment(ApartmentState.STA)]
        public void GivenPrivateSub_DefaultQuickFixRemovesParameter()
        {
            const string inputCode = @"
Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode = @"
Private Sub Foo()
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                new RemoveUnusedParameterQuickFix(vbe.Object, state, new Mock<IMessageBox>().Object, rewritingManager)
                    .Fix(inspectionResults.First(), rewriteSession);

                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }
    }
}
