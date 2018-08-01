using NUnit.Framework;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.Common;
using Rubberduck.VBEditor;
using System.Windows.Forms;

namespace RubberduckTests.AutoComplete
{
    [TestFixture]
    public class CodeStringTests
    {
        [Test]
        public void ToStringIncludesCaretPipe()
        {
            var input = @"foo = MsgBox(|)";
            var sut = new CodeString(input, new Selection(0, input.IndexOf('|')));

            Assert.AreEqual(input, sut.ToString());
        }

        [Test]
        public void CodeExcludesCaretPipe()
        {
            var input = @"foo = MsgBox(|)";
            var expected = @"foo = MsgBox()";
            var sut = new CodeString(input, new Selection(0, input.IndexOf('|')));

            Assert.AreEqual(expected, sut.Code);
        }
    }

    [TestFixture]
    public class SelfClosingPairCompletionTests
    {
        private CodeString Run(SelfClosingPair pair, CodeString original, char input)
        {
            var sut = new SelfClosingPairCompletionService();
            return sut.Execute(pair, original, input);
        }

        private CodeString Run(SelfClosingPair pair, CodeString original, Keys input)
        {
            var sut = new SelfClosingPairCompletionService();
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
)".ToCodeString();
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
    }
}
