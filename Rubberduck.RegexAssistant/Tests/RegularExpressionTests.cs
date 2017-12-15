using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.RegexAssistant.Tests
{
    [TestClass]
    public class RegularExpressionTests
    {
        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseSingleLiteralGroupAsAtomWorks()
        {
            var pattern = "(g){2,4}";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Group("(g)", new Quantifier("{2,4}")), (expression as SingleAtomExpression).Atom);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseCharacterClassAsAtomWorks()
        {
            var pattern = "[abcd]*";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new CharacterClass("[abcd]", new Quantifier("*")), (expression as SingleAtomExpression).Atom);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseLiteralAsAtomWorks()
        {
            var pattern = "a";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Literal("a", Quantifier.None), (expression as SingleAtomExpression).Atom);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseUnicodeEscapeAsAtomWorks()
        {
            var pattern = "\\u1234+";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Literal("\\u1234", new Quantifier("+")), (expression as SingleAtomExpression).Atom);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseHexEscapeSequenceAsAtomWorks()
        {
            var pattern = "\\x12?";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Literal("\\x12", new Quantifier("?")), (expression as SingleAtomExpression).Atom);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseOctalEscapeSequenceAsAtomWorks()
        {
            var pattern = "\\712{2}";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Literal("\\712", new Quantifier("{2}")), (expression as SingleAtomExpression).Atom);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseEscapedLiteralAsAtomWorks()
        {
            var pattern = "\\)";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Literal("\\)", Quantifier.None), (expression as SingleAtomExpression).Atom);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseUnescapedSpecialCharAsAtomFails()
        {
            foreach (var paren in "()[]{}*?+".ToCharArray().Select(c => "" + c))
            {
                var hack = paren;
                Assert.IsFalse(RegularExpression.TryParseAsAtom(ref hack, out var expression));
                Assert.IsNull(expression);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseSimpleLiteralConcatenationAsConcatenatedExpression()
        {
            var expected = new List<IRegularExpression>
            {
                new SingleAtomExpression(new Literal("a", Quantifier.None)),
                new SingleAtomExpression(new Literal("b", Quantifier.None))
            };

            var expression = RegularExpression.Parse("ab");
            Assert.IsInstanceOfType(expression, typeof(ConcatenatedExpression));
            var subexpressions = (expression as ConcatenatedExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (var i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseSimplisticGroupConcatenationAsConcatenatedExpression()
        {
            var expected = new List<IRegularExpression>
            {
                new SingleAtomExpression(new Literal("a", Quantifier.None)),
                new SingleAtomExpression(new Group("(abc)", new Quantifier("{1,4}"))),
                new SingleAtomExpression(new Literal("b", Quantifier.None))
            };

            var expression = RegularExpression.Parse("a(abc){1,4}b");
            Assert.IsInstanceOfType(expression, typeof(ConcatenatedExpression));
            var subexpressions = (expression as ConcatenatedExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (var i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseSimplisticCharacterClassConcatenationAsConcatenatedExpression()
        {
            var expected = new List<IRegularExpression>
            {
                new SingleAtomExpression(new Literal("a", Quantifier.None)),
                new SingleAtomExpression(new CharacterClass("[abc]", new Quantifier("*"))),
                new SingleAtomExpression(new Literal("b", Quantifier.None))
            };

            var expression = RegularExpression.Parse("a[abc]*b");
            Assert.IsInstanceOfType(expression, typeof(ConcatenatedExpression));
            var subexpressions = (expression as ConcatenatedExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (var i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void ParseSimplisticAlternativesExpression()
        {
            var expected = new List<IRegularExpression>
            {
                new SingleAtomExpression(new Literal("a", Quantifier.None)),
                new SingleAtomExpression(new Literal("b", Quantifier.None))
            };

            var expression = RegularExpression.Parse("a|b");
            Assert.IsInstanceOfType(expression, typeof(AlternativesExpression));
            var subexpressions = (expression as AlternativesExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (var i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void CharacterClassIsNotAnAlternativesExpression()
        {
            var expression = RegularExpression.Parse("[a|b]");
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new CharacterClass("[a|b]", Quantifier.None), (expression as SingleAtomExpression).Atom);
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void GroupIsNotAnAlternativesExpression()
        {
            var expression = RegularExpression.Parse("(a|b)");
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Group("(a|b)", Quantifier.None), (expression as SingleAtomExpression).Atom);
        }
    }
}
