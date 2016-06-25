using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.RegexAssistant;
using System;
using System.Linq;

namespace RegexAssistantTests
{
    [TestClass]
    public class LiteralTests
    {
        [TestMethod]
        public void EscapedLiteralTests()
        {
            char[] literals = new char[] { '(', ')', '{', '}', '[', ']', '.', '?', '+', '*' };
            foreach (char literal in literals)
            {
                Literal cut = new Literal("\\" + literal);
                Assert.AreEqual("\\" + literal, cut.Specifier);
            }
        }

        [TestMethod]
        public void EscapeSequences()
        {
            char[] escapes = "sSwWbBdDrnvtf123456789".ToCharArray();
            foreach (char escape in escapes)
            {
                Literal cut = new Literal("\\" + escape);
                Assert.AreEqual("\\" + escape, cut.Specifier);
            }
        }

        [TestMethod]
        public void CodePoints()
        {
            string[] codePoints = { @"\uFFFF", @"\u0000", @"\xFF", @"\x00", @"\777", @"\000" };
            foreach (string codePoint in codePoints)
            {
                Literal cut = new Literal(codePoint);
                Assert.AreEqual(codePoint, cut.Specifier);
            }
        }

        [TestMethod]
        public void SimpleLiterals()
        {
            char[] literals = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!\"§%&/=ß#'°".ToCharArray();
            foreach (char literal in literals)
            {
                Literal cut = new Literal("" + literal);
                Assert.AreEqual("" + literal, cut.Specifier);
            }
        }

        [TestMethod]
        public void EverythingElseBlowsUp()
        {
            char[] allowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!\"§%&/=ß#'°".ToCharArray();
            string[] allowedEscapes = { "(", ")", "{", "}", "[", "]", ".", "?", "+", "*", "$", "^", "uFFFF", "u0000", "xFF", "x00", "777", "000" };
            foreach (string blowup in allowedEscapes.Select(e => "\\"+ e).Concat(allowed.Select(c => ""+c)))
            {
                try
                {
                    Literal cut = new Literal("a"+blowup);
                }
                catch (ArgumentException ex)
                {
                    Assert.IsTrue(true); // Assert.Pass();
                    continue;
                }
                Assert.Fail("Did not blow up when trying to parse {0} as literal", blowup);
            }
        }

        [TestMethod]
        public void SingleEscapedCharsAreNotParsedAsLiteral()
        {
            string[] escapedChars = "(){}[]\\*?+$^".ToCharArray().Select(e => e.ToString()).ToArray();
            foreach (var escape in escapedChars)
            {
                try
                {
                    Literal cut = new Literal(escape);
                }
                catch (ArgumentException ex)
                {
                    Assert.IsTrue(true); // Assert.Pass();
                    continue;
                }
                Assert.Fail("Did not blow up when trying to parse {0} as literal", escape);
            }

        }
    }
}
