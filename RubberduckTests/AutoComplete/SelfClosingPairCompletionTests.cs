using NUnit.Framework;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.VBEditor;

namespace RubberduckTests.AutoComplete
{
    [TestFixture]
    public class SelfClosingPairCompletionTests
    {
        private (string Code, Selection zPosition) Run(SelfClosingPair pair, (string,Selection) original, char input)
        {
            var sut = new SelfClosingPairCompletionService();
            return sut.Execute(pair, original, input);
        }

        [Test]
        public void WhenNextPositionIsClosingChar_ClosingCharMovesSelection()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = pair.ClosingChar;
            var original = @"foo = MsgBox(|)".RemovePseudoCaret();
            var expected = @"foo = MsgBox()|".RemovePseudoCaret();

            var result = Run(pair, original, input);

            Assert.AreEqual(
                expected.Code.InsertPseudoCaret(expected.zPosition).Code, 
                result.Code.InsertPseudoCaret(result.zPosition).Code);
        }

        [Test]
        public void WhenNextPositionIsClosingChar_OpeningCharInsertsNewPair()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = pair.OpeningChar;
            var original = @"foo = MsgBox(|)".RemovePseudoCaret();
            var expected = @"foo = MsgBox((|))".RemovePseudoCaret();

            var result = Run(pair, original, input);

            Assert.AreEqual(
                expected.Code.InsertPseudoCaret(expected.zPosition).Code, 
                result.Code.InsertPseudoCaret(result.zPosition).Code);
        }
    }
}
