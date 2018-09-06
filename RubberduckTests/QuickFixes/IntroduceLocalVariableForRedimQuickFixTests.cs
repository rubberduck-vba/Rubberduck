using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Inspections;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IntroduceLocalVariableForRedimQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void VariantRedimStatement_QuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo()
Redim Foo(1)
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
Dim Foo As Variant
Redim Foo(1)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UndeclaredRedimVariableInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IntroduceLocalVariableForRedimQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }
    }
}
