using NUnit.Framework;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.RegexAssistant.Tests
{
    [TestFixture]
    public class RegularExpressionTests
    {
        [Category("RegexAssistant")]
        [Test]
        public void ParseSingleLiteralGroupAsAtomWorks()
        {
            var pattern = "(g){2,4}";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOf(typeof(SingleAtomExpression), expression);
            Assert.AreEqual(new Group("(g)", new Quantifier("{2,4}")), (expression as SingleAtomExpression).Atom);
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseCharacterClassAsAtomWorks()
        {
            var pattern = "[abcd]*";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOf(typeof(SingleAtomExpression), expression);
            Assert.AreEqual(new CharacterClass("[abcd]", new Quantifier("*")), (expression as SingleAtomExpression).Atom);
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseLiteralAsAtomWorks()
        {
            var pattern = "a";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOf(typeof(SingleAtomExpression), expression);
            Assert.AreEqual(new Literal("a", Quantifier.None), (expression as SingleAtomExpression).Atom);
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseUnicodeEscapeAsAtomWorks()
        {
            var pattern = "\\u1234+";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOf(typeof(SingleAtomExpression), expression);
            Assert.AreEqual(new Literal("\\u1234", new Quantifier("+")), (expression as SingleAtomExpression).Atom);
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseHexEscapeSequenceAsAtomWorks()
        {
            var pattern = "\\x12?";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOf(typeof(SingleAtomExpression), expression);
            Assert.AreEqual(new Literal("\\x12", new Quantifier("?")), (expression as SingleAtomExpression).Atom);
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseOctalEscapeSequenceAsAtomWorks()
        {
            var pattern = "\\712{2}";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOf(typeof(SingleAtomExpression), expression);
            Assert.AreEqual(new Literal("\\712", new Quantifier("{2}")), (expression as SingleAtomExpression).Atom);
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseEscapedLiteralAsAtomWorks()
        {
            var pattern = "\\)";
            RegularExpression.TryParseAsAtom(ref pattern, out var expression);
            Assert.IsInstanceOf(typeof(SingleAtomExpression), expression);
            Assert.AreEqual(new Literal("\\)", Quantifier.None), (expression as SingleAtomExpression).Atom);
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseUnescapedSpecialCharAsAtomFails()
        {
            foreach (var paren in "()[]{}*?+".ToCharArray().Select(c => "" + c))
            {
                var hack = paren;
                Assert.IsFalse(RegularExpression.TryParseAsAtom(ref hack, out var expression));
                Assert.IsNull(expression);
            }
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseSimpleLiteralConcatenationAsConcatenatedExpression()
        {
            var expected = new List<IRegularExpression>
            {
                new SingleAtomExpression(new Literal("a", Quantifier.None)),
                new SingleAtomExpression(new Literal("b", Quantifier.None))
            };

            var expression = RegularExpression.Parse("ab");
            Assert.IsInstanceOf(typeof(ConcatenatedExpression), expression);
            var subexpressions = (expression as ConcatenatedExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (var i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseSimplisticGroupConcatenationAsConcatenatedExpression()
        {
            var expected = new List<IRegularExpression>
            {
                new SingleAtomExpression(new Literal("a", Quantifier.None)),
                new SingleAtomExpression(new Group("(abc)", new Quantifier("{1,4}"))),
                new SingleAtomExpression(new Literal("b", Quantifier.None))
            };

            var expression = RegularExpression.Parse("a(abc){1,4}b");
            Assert.IsInstanceOf(typeof(ConcatenatedExpression), expression);
            var subexpressions = (expression as ConcatenatedExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (var i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseSimplisticCharacterClassConcatenationAsConcatenatedExpression()
        {
            var expected = new List<IRegularExpression>
            {
                new SingleAtomExpression(new Literal("a", Quantifier.None)),
                new SingleAtomExpression(new CharacterClass("[abc]", new Quantifier("*"))),
                new SingleAtomExpression(new Literal("b", Quantifier.None))
            };

            var expression = RegularExpression.Parse("a[abc]*b");
            Assert.IsInstanceOf(typeof(ConcatenatedExpression), expression);
            var subexpressions = (expression as ConcatenatedExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (var i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [Category("RegexAssistant")]
        [Test]
        public void ParseSimplisticAlternativesExpression()
        {
            var expected = new List<IRegularExpression>
            {
                new SingleAtomExpression(new Literal("a", Quantifier.None)),
                new SingleAtomExpression(new Literal("b", Quantifier.None))
            };

            var expression = RegularExpression.Parse("a|b");
            Assert.IsInstanceOf(typeof(AlternativesExpression), expression);
            var subexpressions = (expression as AlternativesExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (var i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [Category("RegexAssistant")]
        [Test]
        public void CharacterClassIsNotAnAlternativesExpression()
        {
            var expression = RegularExpression.Parse("[a|b]");
            Assert.IsInstanceOf(typeof(SingleAtomExpression), expression);
            Assert.AreEqual(new CharacterClass("[a|b]", Quantifier.None), (expression as SingleAtomExpression).Atom);
        }

        [Category("RegexAssistant")]
        [Test]
        public void GroupIsNotAnAlternativesExpression()
        {
            var expression = RegularExpression.Parse("(a|b)");
            Assert.IsInstanceOf(typeof(SingleAtomExpression), expression);
            Assert.AreEqual(new Group("(a|b)", Quantifier.None), (expression as SingleAtomExpression).Atom);
        }
    }
}
