using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Rubberduck.RegexAssistant.Tests
{
    [TestClass]
    public class CharacterClassTests
    {
        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void InvertedCharacterClass()
        {
            CharacterClass cut = new CharacterClass("[^ ]", Quantifier.None);
            Assert.IsTrue(cut.InverseMatching);
            List<string> expectedSpecifiers = new List<string>();
            expectedSpecifiers.Add(" ");

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (int i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void SimpleCharacterRange()
        {
            CharacterClass cut = new CharacterClass("[a-z]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            List<string> expectedSpecifiers = new List<string>();
            expectedSpecifiers.Add("a-z");

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (int i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void UnicodeCharacterRange()
        {
            CharacterClass cut = new CharacterClass(@"[\u00A2-\uFFFF]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            List<string> expectedSpecifiers = new List<string>();
            expectedSpecifiers.Add(@"\u00A2-\uFFFF");

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (int i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void OctalCharacterRange()
        {
            CharacterClass cut = new CharacterClass(@"[\011-\777]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            List<string> expectedSpecifiers = new List<string>();
            expectedSpecifiers.Add(@"\011-\777");

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (int i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void HexadecimalCharacterRange()
        {
            CharacterClass cut = new CharacterClass(@"[\x00-\xFF]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            List<string> expectedSpecifiers = new List<string>();
            expectedSpecifiers.Add(@"\x00-\xFF");

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (int i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void MixedCharacterRanges()
        {
            CharacterClass cut = new CharacterClass(@"[\x00-\777\u001A-Z]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            List<string> expectedSpecifiers = new List<string>();
            expectedSpecifiers.Add(@"\x00-\777");
            expectedSpecifiers.Add(@"\u001A-Z");

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (int i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void RangeFailureWithCharacterClass()
        {
            foreach (string charClass in new string[]{ @"\D", @"\d", @"\s", @"\S", @"\w", @"\W" }){
                CharacterClass cut = new CharacterClass(string.Format("[{0}-F]", charClass), Quantifier.None);
                Assert.IsFalse(cut.InverseMatching);
                List<string> expectedSpecifiers = new List<string>();
                expectedSpecifiers.Add(charClass);
                expectedSpecifiers.Add(@"-");
                expectedSpecifiers.Add(@"F");

                Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
                for (int i = 0; i < expectedSpecifiers.Count; i++)
                {
                    Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
                }
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void EscapedLiteralRanges()
        {
            foreach (string escapedLiteral in new string[] { @"\.", @"\[", @"\]" })
            {
                CharacterClass cut = new CharacterClass(string.Format("[{0}-F]", escapedLiteral), Quantifier.None);
                Assert.IsFalse(cut.InverseMatching);
                List<string> expectedSpecifiers = new List<string>();
                expectedSpecifiers.Add(string.Format("{0}-F",escapedLiteral));

                Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
                for (int i = 0; i < expectedSpecifiers.Count; i++)
                {
                    Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
                }
                // invert this
                cut = new CharacterClass(string.Format("[F-{0}]", escapedLiteral), Quantifier.None);
                Assert.IsFalse(cut.InverseMatching);
                expectedSpecifiers.Clear();
                expectedSpecifiers.Add(string.Format("F-{0}", escapedLiteral));

                Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
                for (int i = 0; i < expectedSpecifiers.Count; i++)
                {
                    Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
                }
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void SkipsIncorrectlyEscapedLiterals()
        {
            foreach (string escapedLiteral in new string[] { @"\(", @"\)", @"\{", @"\}", @"\|", @"\?", @"\*" })
            {
                CharacterClass cut = new CharacterClass(string.Format("[{0}-F]", escapedLiteral), Quantifier.None);
                Assert.IsFalse(cut.InverseMatching);
                List<string> expectedSpecifiers = new List<string>();
                expectedSpecifiers.Add(string.Format("{0}-F", escapedLiteral.Substring(1)));

                Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
                for (int i = 0; i < expectedSpecifiers.Count; i++)
                {
                    Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
                }
                // inverted doesn't need to behave the same, because VBA blows up for ranges like R-\(

            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void IncorrectlyEscapedRangeTargetLiteralsBlowUp()
        {
            foreach (string escapedLiteral in new string[] { @"\(", @"\)", @"\{", @"\}", @"\|", @"\?", @"\*" })
            {
                try
                {
                    CharacterClass cut = new CharacterClass(string.Format("[F-{0}]", escapedLiteral), Quantifier.None);
                }
#pragma warning disable CS0168 // Variable is declared but never used
                catch (ArgumentException ex)
#pragma warning restore CS0168 // Variable is declared but never used
                {
                    continue;
                }
                Assert.Fail("Incorrectly allowed character range with {0} as target", escapedLiteral);
            }

        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void IgnoresBackreferenceSpecifiers()
        {
            CharacterClass cut = new CharacterClass(@"[\1]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);

            List<string> expectedSpecifiers = new List<string>();
            expectedSpecifiers.Add("1");
            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (int i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }
    }
}
