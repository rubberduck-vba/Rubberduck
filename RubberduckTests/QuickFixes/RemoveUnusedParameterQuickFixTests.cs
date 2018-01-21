using System.Linq;
using System.Threading;
using NUnit.Framework;
using Moq;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.UI;
using RubberduckTests.Mocks;

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
            using(var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new RemoveUnusedParameterQuickFix(vbe.Object, state, new Mock<IMessageBox>().Object).Fix(
                    inspectionResults.First());
                Assert.AreEqual(expectedCode, component.CodeModule.Content());
            }
        }
    }
}
