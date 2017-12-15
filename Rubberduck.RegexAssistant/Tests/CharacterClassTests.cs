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
            var cut = new CharacterClass("[^ ]", Quantifier.None);
            Assert.IsTrue(cut.InverseMatching);
            var expectedSpecifiers = new List<string> { " " };

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (var i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void SimpleCharacterRange()
        {
            var cut = new CharacterClass("[a-z]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            var expectedSpecifiers = new List<string> { "a-z" };

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (var i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void UnicodeCharacterRange()
        {
            var cut = new CharacterClass(@"[\u00A2-\uFFFF]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            var expectedSpecifiers = new List<string> { @"\u00A2-\uFFFF" };

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (var i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void OctalCharacterRange()
        {
            var cut = new CharacterClass(@"[\011-\777]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            var expectedSpecifiers = new List<string> { @"\011-\777" };

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (var i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void HexadecimalCharacterRange()
        {
            var cut = new CharacterClass(@"[\x00-\xFF]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            var expectedSpecifiers = new List<string> { @"\x00-\xFF" };

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (var i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void MixedCharacterRanges()
        {
            var cut = new CharacterClass(@"[\x00-\777\u001A-Z]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);
            var expectedSpecifiers = new List<string>
            {
                @"\x00-\777",
                @"\u001A-Z"
            };

            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (var i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void RangeFailureWithCharacterClass()
        {
            foreach (var charClass in new[]{ @"\D", @"\d", @"\s", @"\S", @"\w", @"\W" }){
                var cut = new CharacterClass($"[{charClass}-F]", Quantifier.None);
                Assert.IsFalse(cut.InverseMatching);
                var expectedSpecifiers = new List<string>
                {
                    charClass,
                    @"-",
                    @"F"
                };

                Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
                for (var i = 0; i < expectedSpecifiers.Count; i++)
                {
                    Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
                }
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void EscapedLiteralRanges()
        {
            foreach (var escapedLiteral in new[] { @"\.", @"\[", @"\]" })
            {
                var cut = new CharacterClass($"[{escapedLiteral}-F]", Quantifier.None);
                Assert.IsFalse(cut.InverseMatching);
                var expectedSpecifiers = new List<string> {$"{escapedLiteral}-F"};

                Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
                for (var i = 0; i < expectedSpecifiers.Count; i++)
                {
                    Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
                }
                // invert this
                cut = new CharacterClass($"[F-{escapedLiteral}]", Quantifier.None);
                Assert.IsFalse(cut.InverseMatching);
                expectedSpecifiers.Clear();
                expectedSpecifiers.Add($"F-{escapedLiteral}");

                Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
                for (var i = 0; i < expectedSpecifiers.Count; i++)
                {
                    Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
                }
            }
        }

        [TestCategory("RegexAssistant")]
        [TestMethod]
        public void SkipsIncorrectlyEscapedLiterals()
        {
            foreach (var escapedLiteral in new[] { @"\(", @"\)", @"\{", @"\}", @"\|", @"\?", @"\*" })
            {
                var cut = new CharacterClass($"[{escapedLiteral}-F]", Quantifier.None);
                Assert.IsFalse(cut.InverseMatching);
                var expectedSpecifiers = new List<string> {$"{escapedLiteral.Substring(1)}-F"};

                Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
                for (var i = 0; i < expectedSpecifiers.Count; i++)
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
            foreach (var escapedLiteral in new[] { @"\(", @"\)", @"\{", @"\}", @"\|", @"\?", @"\*" })
            {
                try
                {
                    var cut = new CharacterClass($"[F-{escapedLiteral}]", Quantifier.None);
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
            var cut = new CharacterClass(@"[\1]", Quantifier.None);
            Assert.IsFalse(cut.InverseMatching);

            var expectedSpecifiers = new List<string> {"1"};
            Assert.AreEqual(expectedSpecifiers.Count, cut.CharacterSpecifiers.Count);
            for (var i = 0; i < expectedSpecifiers.Count; i++)
            {
                Assert.AreEqual(expectedSpecifiers[i], cut.CharacterSpecifiers[i]);
            }
        }
    }
}
