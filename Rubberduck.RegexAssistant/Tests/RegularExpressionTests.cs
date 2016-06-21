using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.RegexAssistant.Tests
{
    [TestClass]
    public class RegularExpressionTests
    {
        [TestMethod]
        public void ParseSingleLiteralGroupAsAtomWorks()
        {
            IRegularExpression expression;
            string pattern = "(g){2,4}";
            RegularExpression.TryParseAsAtom(ref pattern, out expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Quantifier("{2,4}"), expression.Quantifier);
            Assert.AreEqual(new Group("(g)"), (expression as SingleAtomExpression).Atom);
        }

        [TestMethod]
        public void ParseCharacterClassAsAtomWorks()
        {
            IRegularExpression expression;
            string pattern = "[abcd]*";
            RegularExpression.TryParseAsAtom(ref pattern, out expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Quantifier("*"), expression.Quantifier);
            Assert.AreEqual(new CharacterClass("[abcd]"), (expression as SingleAtomExpression).Atom);
        }

        [TestMethod]
        public void ParseLiteralAsAtomWorks()
        {
            IRegularExpression expression;
            string pattern = "a";
            RegularExpression.TryParseAsAtom(ref pattern, out expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Quantifier(""), expression.Quantifier);
            Assert.AreEqual(new Literal("a"), (expression as SingleAtomExpression).Atom);
        }

        [TestMethod]
        public void ParseUnicodeEscapeAsAtomWorks()
        {
            IRegularExpression expression;
            string pattern = "\\u1234+";
            RegularExpression.TryParseAsAtom(ref pattern, out expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Quantifier("+"), expression.Quantifier);
            Assert.AreEqual(new Literal("\\u1234"), (expression as SingleAtomExpression).Atom);
        }

        [TestMethod]
        public void ParseHexEscapeSequenceAsAtomWorks()
        {
            IRegularExpression expression;
            string pattern = "\\x12?";
            RegularExpression.TryParseAsAtom(ref pattern, out expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Quantifier("?"), expression.Quantifier);
            Assert.AreEqual(new Literal("\\x12"), (expression as SingleAtomExpression).Atom);
        }

        [TestMethod]
        public void ParseOctalEscapeSequenceAsAtomWorks()
        {
            IRegularExpression expression;
            string pattern = "\\712{2}";
            RegularExpression.TryParseAsAtom(ref pattern, out expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Quantifier("{2}"), expression.Quantifier);
            Assert.AreEqual(new Literal("\\712"), (expression as SingleAtomExpression).Atom);
        }

        [TestMethod]
        public void ParseEscapedLiteralAsAtomWorks()
        {
            IRegularExpression expression;
            string pattern = "\\)";
            RegularExpression.TryParseAsAtom(ref pattern, out expression);
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Quantifier(""), expression.Quantifier);
            Assert.AreEqual(new Literal("\\)"), (expression as SingleAtomExpression).Atom);
        }

        [TestMethod]
        public void ParseUnescapedSpecialCharAsAtomFails()
        {
            foreach (string paren in "()[]{}*?+".ToCharArray().Select(c => "" + c))
            {
                IRegularExpression expression;
                string hack = paren;
                Assert.IsFalse(RegularExpression.TryParseAsAtom(ref hack, out expression));
                Assert.IsNull(expression);
            }
        }

        [TestMethod]
        public void ParseSimpleLiteralConcatenationAsConcatenatedExpression()
        {
            List<IRegularExpression> expected = new List<IRegularExpression>();
            expected.Add(new SingleAtomExpression(new Literal("a"), new Quantifier("")));
            expected.Add(new SingleAtomExpression(new Literal("b"), new Quantifier("")));

            IRegularExpression expression = RegularExpression.Parse("ab");
            Assert.IsInstanceOfType(expression, typeof(ConcatenatedExpression));
            var subexpressions = (expression as ConcatenatedExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [TestMethod]
        public void ParseSimplisticGroupConcatenationAsConcatenatedExpression()
        {
            List<IRegularExpression> expected = new List<IRegularExpression>();
            expected.Add(new SingleAtomExpression(new Literal("a"), new Quantifier("")));
            expected.Add(new SingleAtomExpression(new Group("(abc)"), new Quantifier("{1,4}")));
            expected.Add(new SingleAtomExpression(new Literal("b"), new Quantifier("")));

            IRegularExpression expression = RegularExpression.Parse("a(abc){1,4}b");
            Assert.IsInstanceOfType(expression, typeof(ConcatenatedExpression));
            var subexpressions = (expression as ConcatenatedExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [TestMethod]
        public void ParseSimplisticCharacterClassConcatenationAsConcatenatedExpression()
        {
            List<IRegularExpression> expected = new List<IRegularExpression>();
            expected.Add(new SingleAtomExpression(new Literal("a"), new Quantifier("")));
            expected.Add(new SingleAtomExpression(new CharacterClass("[abc]"), new Quantifier("*")));
            expected.Add(new SingleAtomExpression(new Literal("b"), new Quantifier("")));

            IRegularExpression expression = RegularExpression.Parse("a[abc]*b");
            Assert.IsInstanceOfType(expression, typeof(ConcatenatedExpression));
            var subexpressions = (expression as ConcatenatedExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [TestMethod]
        public void ParseSimplisticAlternativesExpression()
        {
            List<IRegularExpression> expected = new List<IRegularExpression>();
            expected.Add(new SingleAtomExpression(new Literal("a"), new Quantifier("")));
            expected.Add(new SingleAtomExpression(new Literal("b"), new Quantifier("")));

            IRegularExpression expression = RegularExpression.Parse("a|b");
            Assert.IsInstanceOfType(expression, typeof(AlternativesExpression));
            var subexpressions = (expression as AlternativesExpression).Subexpressions;
            Assert.AreEqual(expected.Count, subexpressions.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.AreEqual(expected[i], subexpressions[i]);
            }
        }

        [TestMethod]
        public void CharacterClassIsNotAnAlternativesExpression()
        {
            IRegularExpression expression = RegularExpression.Parse("[a|b]");
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new CharacterClass("[a|b]"), (expression as SingleAtomExpression).Atom);
        }

        [TestMethod]
        public void GroupIsNotAnAlternativesExpression()
        {
            IRegularExpression expression = RegularExpression.Parse("(a|b)");
            Assert.IsInstanceOfType(expression, typeof(SingleAtomExpression));
            Assert.AreEqual(new Group("(a|b)"), (expression as SingleAtomExpression).Atom);
        }
    }
}
