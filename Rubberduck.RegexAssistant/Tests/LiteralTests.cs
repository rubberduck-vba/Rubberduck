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
            char[] literals = new char[] { '(', ')', '{', '}', '[', ']', '.', '?', '+', '*' };
            foreach (char literal in literals)
            {
                Literal cut = new Literal("\\" + literal, Quantifier.None);
                Assert.AreEqual("\\" + literal, cut.Specifier);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void EscapeSequences()
        {
            char[] escapes = "sSwWbBdDrnvtf123456789".ToCharArray();
            foreach (char escape in escapes)
            {
                Literal cut = new Literal("\\" + escape, Quantifier.None);
                Assert.AreEqual("\\" + escape, cut.Specifier);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void CodePoints()
        {
            string[] codePoints = { @"\uFFFF", @"\u0000", @"\xFF", @"\x00", @"\777", @"\000" };
            foreach (string codePoint in codePoints)
            {
                Literal cut = new Literal(codePoint, Quantifier.None);
                Assert.AreEqual(codePoint, cut.Specifier);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void SimpleLiterals()
        {
            char[] literals = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!\"§%&/=ß#'°".ToCharArray();
            foreach (char literal in literals)
            {
                Literal cut = new Literal("" + literal, Quantifier.None);
                Assert.AreEqual("" + literal, cut.Specifier);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void EverythingElseBlowsUp()
        {
            char[] allowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!\"§%&/=ß#'°".ToCharArray();
            string[] allowedEscapes = { "(", ")", "{", "}", "[", "]", ".", "?", "+", "*", "$", "^", "uFFFF", "u0000", "xFF", "x00", "777", "000" };
            foreach (string blowup in allowedEscapes.Select(e => "\\"+ e).Concat(allowed.Select(c => ""+c)))
            {
                try
                {
                    Literal cut = new Literal("a"+blowup, Quantifier.None);
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
            string[] escapedChars = "(){}[]\\*?+$^".ToCharArray().Select(e => e.ToString()).ToArray();
            foreach (var escape in escapedChars)
            {
                try
                {
                    Literal cut = new Literal(escape, Quantifier.None);
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
