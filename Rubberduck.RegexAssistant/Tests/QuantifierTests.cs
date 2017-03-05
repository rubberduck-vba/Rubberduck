using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Rubberduck.RegexAssistant.Tests
{
    [TestClass]
    public class QuantifierTests
    {
        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void AsteriskQuantifier()
        {
            Quantifier cut = new Quantifier("*");
            Assert.AreEqual(int.MaxValue, cut.MaximumMatches);
            Assert.AreEqual(0, cut.MinimumMatches);
            Assert.AreEqual(QuantifierKind.Wildcard, cut.Kind);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void QuestionMarkQuantifier()
        {
            Quantifier cut = new Quantifier("?");
            Assert.AreEqual(1, cut.MaximumMatches);
            Assert.AreEqual(0, cut.MinimumMatches);
            Assert.AreEqual(QuantifierKind.Wildcard, cut.Kind);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void PlusQuantifier()
        {
            Quantifier cut = new Quantifier("+");
            Assert.AreEqual(int.MaxValue, cut.MaximumMatches);
            Assert.AreEqual(1, cut.MinimumMatches);
            Assert.AreEqual(QuantifierKind.Wildcard, cut.Kind);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ExactQuantifier()
        {
            Quantifier cut = new Quantifier("{5}");
            Assert.AreEqual(5, cut.MaximumMatches);
            Assert.AreEqual(5, cut.MinimumMatches);
            Assert.AreEqual(QuantifierKind.Expression, cut.Kind);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void FullRangeQuantifier()
        {
            Quantifier cut = new Quantifier("{2,5}");
            Assert.AreEqual(2, cut.MinimumMatches);
            Assert.AreEqual(5, cut.MaximumMatches);
            Assert.AreEqual(QuantifierKind.Expression, cut.Kind);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void OpenRangeQuantifier()
        {
            Quantifier cut = new Quantifier("{3,}");
            Assert.AreEqual(3, cut.MinimumMatches);
            Assert.AreEqual(int.MaxValue, cut.MaximumMatches);
            Assert.AreEqual(QuantifierKind.Expression, cut.Kind);

        }
    }
}
