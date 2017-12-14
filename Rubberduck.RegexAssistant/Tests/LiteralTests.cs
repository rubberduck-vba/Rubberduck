using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Rubberduck.RegexAssistant.Tests
{
    [TestClass]
    public class LiteralTests
    {
        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void EscapedLiteralTests()
        {
            var literals = new[] { '(', ')', '{', '}', '[', ']', '.', '?', '+', '*' };
            foreach (var literal in literals)
            {
                var cut = new Literal("\\" + literal, Quantifier.None);
                Assert.AreEqual("\\" + literal, cut.Specifier);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void EscapeSequences()
        {
            var escapes = "sSwWbBdDrnvtf123456789".ToCharArray();
            foreach (var escape in escapes)
            {
                var cut = new Literal("\\" + escape, Quantifier.None);
                Assert.AreEqual("\\" + escape, cut.Specifier);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void CodePoints()
        {
            string[] codePoints = { @"\uFFFF", @"\u0000", @"\xFF", @"\x00", @"\777", @"\000" };
            foreach (var codePoint in codePoints)
            {
                var cut = new Literal(codePoint, Quantifier.None);
                Assert.AreEqual(codePoint, cut.Specifier);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void SimpleLiterals()
        {
            var literals = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!\"§%&/=ß#'°".ToCharArray();
            foreach (var literal in literals)
            {
                var cut = new Literal("" + literal, Quantifier.None);
                Assert.AreEqual("" + literal, cut.Specifier);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void EverythingElseBlowsUp()
        {
            var allowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!\"§%&/=ß#'°".ToCharArray();
            string[] allowedEscapes = { "(", ")", "{", "}", "[", "]", ".", "?", "+", "*", "$", "^", "uFFFF", "u0000", "xFF", "x00", "777", "000" };
            foreach (var blowup in allowedEscapes.Select(e => "\\"+ e).Concat(allowed.Select(c => ""+c)))
            {
                try
                {
                    var cut = new Literal("a"+blowup, Quantifier.None);
                }
#pragma warning disable CS0168 // Variable is declared but never used
                catch (ArgumentException ex)
#pragma warning restore CS0168 // Variable is declared but never used
                {
                    Assert.IsTrue(true); // Assert.Pass();
                    continue;
                }
                Assert.Fail("Did not blow up when trying to parse {0} as literal", blowup);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void SingleEscapedCharsAreNotParsedAsLiteral()
        {
            var escapedChars = "(){}[]\\*?+$^".ToCharArray().Select(e => e.ToString()).ToArray();
            foreach (var escape in escapedChars)
            {
                try
                {
                    var cut = new Literal(escape, Quantifier.None);
                }
#pragma warning disable CS0168 // Variable is declared but never used
                catch (ArgumentException ex)
#pragma warning restore CS0168 // Variable is declared but never used
                {
                    Assert.IsTrue(true); // Assert.Pass();
                    continue;
                }
                Assert.Fail("Did not blow up when trying to parse {0} as literal", escape);
            }

        }
    }
}
