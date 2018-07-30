using NUnit.Framework;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.VBEditor;

namespace RubberduckTests.AutoComplete
{
    [TestFixture]
    public class SelfClosingPairCompletionTests
    {
        private CodeString Run(SelfClosingPair pair, CodeString original, char input)
        {
            var sut = new SelfClosingPairCompletionService();
            var result = sut.Execute(pair, (original.Code, original.CaretPosition), input);
            return new CodeString(result.Code, result.CaretPosition);
        }

        [Test]
        public void WhenNextPositionIsClosingChar_ClosingCharMovesSelection()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = pair.ClosingChar;
            var original = @"foo = MsgBox(|)".RemovePseudoCaret();
            var expected = @"foo = MsgBox()|".RemovePseudoCaret();

            var result = Run(pair, original, input);

            Assert.AreEqual(expected, result);
        }

        [Test]
        public void WhenNextPositionIsClosingChar_OpeningCharInsertsNewPair()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = pair.OpeningChar;
            var original = @"foo = MsgBox(|)".RemovePseudoCaret();
            var expected = @"foo = MsgBox((|))".RemovePseudoCaret();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }
    }
}
