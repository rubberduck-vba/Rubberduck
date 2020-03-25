using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class QuickFixBaseTests
    {
        [Test]
        [Category(nameof(QuickFixes))]
        public void QuickFixBase_Register()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var quickFix = new RemoveCommentQuickFix();
                quickFix.RegisterInspections(typeof(EmptyStringLiteralInspection));

                Assert.IsTrue(quickFix.SupportedInspections.Contains(typeof(EmptyStringLiteralInspection)));
            }
        }

        [Test]
        [Category(nameof(QuickFixes))]
        public void QuickFixBase_Unregister()
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Empty, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var quickFix = new RemoveCommentQuickFix();
                quickFix.RemoveInspections(quickFix.SupportedInspections.ToArray());

                Assert.IsFalse(quickFix.SupportedInspections.Any());
            }
        }
    }
}