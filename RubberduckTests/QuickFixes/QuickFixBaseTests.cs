using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class QuickFixBaseTests
    {
        [TestMethod]
        [TestCategory(nameof(QuickFixes))]
        public void QuickFixBase_Register()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var quickFix = new RemoveCommentQuickFix(state);
                quickFix.RegisterInspections(typeof(EmptyStringLiteralInspection));

                Assert.IsTrue(quickFix.SupportedInspections.Contains(typeof(EmptyStringLiteralInspection)));
            }
        }

        [TestMethod]
        [TestCategory(nameof(QuickFixes))]
        public void QuickFixBase_Unregister()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var quickFix = new RemoveCommentQuickFix(state);
                quickFix.RemoveInspections(quickFix.SupportedInspections.ToArray());

                Assert.IsFalse(quickFix.SupportedInspections.Any());
            }
        }
    }
}