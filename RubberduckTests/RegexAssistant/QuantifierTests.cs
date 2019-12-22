using NUnit.Framework;

namespace Rubberduck.RegexAssistant.Tests
{
    [TestFixture]
    [Category("RegexAssistant")]
    public class QuantifierTests
    {

        [Test]
        public void AsteriskQuantifier()
        {
            var cut = new Quantifier("*");
            Assert.AreEqual(int.MaxValue, cut.MaximumMatches);
            Assert.AreEqual(0, cut.MinimumMatches);
            Assert.AreEqual(QuantifierKind.Wildcard, cut.Kind);
        }


        [Test]
        public void QuestionMarkQuantifier()
        {
            var cut = new Quantifier("?");
            Assert.AreEqual(1, cut.MaximumMatches);
            Assert.AreEqual(0, cut.MinimumMatches);
            Assert.AreEqual(QuantifierKind.Wildcard, cut.Kind);
        }


        [Test]
        public void PlusQuantifier()
        {
            var cut = new Quantifier("+");
            Assert.AreEqual(int.MaxValue, cut.MaximumMatches);
            Assert.AreEqual(1, cut.MinimumMatches);
            Assert.AreEqual(QuantifierKind.Wildcard, cut.Kind);
        }


        [Test]
        public void ExactQuantifier()
        {
            var cut = new Quantifier("{5}");
            Assert.AreEqual(5, cut.MaximumMatches);
            Assert.AreEqual(5, cut.MinimumMatches);
            Assert.AreEqual(QuantifierKind.Expression, cut.Kind);
        }


        [Test]
        public void FullRangeQuantifier()
        {
            var cut = new Quantifier("{2,5}");
            Assert.AreEqual(2, cut.MinimumMatches);
            Assert.AreEqual(5, cut.MaximumMatches);
            Assert.AreEqual(QuantifierKind.Expression, cut.Kind);
        }


        [Test]
        public void OpenRangeQuantifier()
        {
            var cut = new Quantifier("{3,}");
            Assert.AreEqual(3, cut.MinimumMatches);
            Assert.AreEqual(int.MaxValue, cut.MaximumMatches);
            Assert.AreEqual(QuantifierKind.Expression, cut.Kind);
        }
    }
}
