using NUnit.Framework;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.Common;
using System.Windows.Forms;

namespace RubberduckTests.AutoComplete
{

    [TestFixture]
    public class SelfClosingPairCompletionTests
    {
        private CodeString Run(SelfClosingPair pair, CodeString original, char input)
        {
            var sut = new SelfClosingPairCompletionService(null);
            return sut.Execute(pair, original, input);
        }

        private CodeString Run(SelfClosingPair pair, CodeString original, Keys input)
        { 
            var sut = new SelfClosingPairCompletionService(null);
            return sut.Execute(pair, original, input);
        }

        [Test]
        public void WhenNextPositionIsClosingChar_ClosingCharMovesSelection()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = pair.ClosingChar;
            var original = @"foo = MsgBox(|)".ToCodeString();
            var expected = @"foo = MsgBox()|".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeletingOpeningCharRemovesPairedClosingChar_Parens()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.Back;
            var original = @"foo = (|2 + 2)".ToCodeString();
            var expected = @"foo = |2 + 2".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeletingOpeningCharRemovesPairedClosingChar_StringDelimiter()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = @"foo = ""|2 + 2""".ToCodeString();
            var expected = @"foo = |2 + 2".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeletingOpeningCharRemovesPairedClosingChar_GivenOnlyThatOnTheLine()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = $"{pair.OpeningChar}{pair.ClosingChar}".ToCodeString();
            var expected = string.Empty.ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeletingOpeningCharRemovesPairedClosingChar_NestedParens()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.Back;
            var original = @"foo = ((|2 + 2) + 42)".ToCodeString();
            var expected = @"foo = (|2 + 2 + 42)".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeletingOpeningCharRemovesPairedClosingChar_NestedParensMultiline()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.Back;
            var original = @"
foo = (| _
    (2 + 2) + 42 _
)
".ToCodeString();
            var expected = @"
foo = | _
    (2 + 2) + 42 _

".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void WhenNextPositionIsClosingChar_OpeningCharInsertsNewPair()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = pair.OpeningChar;
            var original = @"foo = MsgBox(|)".ToCodeString();
            var expected = @"foo = MsgBox((|))".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void WhenCaretBetweenOpeningAndClosingChars_BackspaceRemovesBoth()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.Back;
            var original = @"foo = MsgBox(|)".ToCodeString();
            var expected = @"foo = MsgBox|".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void WhenCaretBetweenOpeningAndClosingChars_BackspaceRemovesBoth_Indented()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = @"    foo = ""|""".ToCodeString();
            var expected = @"    foo = |".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeleteKey_ReturnsDefault()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.A;
            var original = @"MsgBox (|)".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsTrue(result == default);
        }

        [Test]
        public void UnhandledKey_ReturnsDefault()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.A;
            var original = @"MsgBox |".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsTrue(result == default);
        }
    }
}
