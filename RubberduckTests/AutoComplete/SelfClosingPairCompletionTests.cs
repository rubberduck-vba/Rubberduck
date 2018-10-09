using NUnit.Framework;
using System.Windows.Forms;
using Rubberduck.AutoComplete.Service;
using Rubberduck.VBEditor;

namespace RubberduckTests.AutoComplete
{
    [TestFixture]
    public class SelfClosingPairCompletionTests
    {
        private TestCodeString Run(SelfClosingPair pair, CodeString original, char input)
        {
            var sut = new SelfClosingPairCompletionService(null);
            var result = sut.Execute(pair, original, input);
            return result != null ? new TestCodeString(result) : null;
        }

        private TestCodeString Run(SelfClosingPair pair, CodeString original, Keys input)
        {
            var sut = new SelfClosingPairCompletionService(null);
            var result = sut.Execute(pair, original, input);
            return result != null ? new TestCodeString(result) : null;
        }

        [Test]
        public void PlacesCaretBetweenOpeningAndClosingChars()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = pair.OpeningChar;
            var original = "foo = MsgBox |".ToCodeString();
            var expected = "foo = MsgBox \"|\"".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void PlacesCaretBetweenOpeningAndClosingChars_PreservesPosition()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = pair.OpeningChar;
            var original = "foo = |".ToCodeString();
            var expected = "foo = (|)".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
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
            var original = $"{pair.OpeningChar}|{pair.ClosingChar}".ToCodeString();
            var expected = string.Empty;

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result?.Code);
        }

        [Test]
        public void BackspacingInsideComment_BailsOut()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.Back;
            var original = "' _\r\n    (|)".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsNull(result);
        }

        [Test]
        public void CanTypeClosingChar()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = pair.ClosingChar;
            var original = "foo = |".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsNull(result);
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
        public void DeletingOpeningChar_CallStmtArgList()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.Back;
            var original = "foo = Call xy(|z)".ToCodeString();
            var expected = "foo = Call xy|z".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeletingOpeningCharRemovesPairedClosingChar_NestedParensMultiline()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.Back;
            var original = @"foo = (| _
    (2 + 2) + 42)
".ToCodeString();
            var expected = @"foo = | _
    (2 + 2) + 42".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test] // fixme: passes, but in-editor behavior seems different.
        public void DeletingPairInLogicalLine_SelectionRemainsOnThatLineIfNonEmpty()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = @"foo = ""abc"" & _
      ""|"" & ""a""".ToCodeString();
            var expected = @"foo = ""abc"" & _
      | & ""a""".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeletingMatchingPair_RemovesTrailingEmptyContinuatedLine()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.Back;
            var original = @"foo = (| _
    (2 + 2) + 42 _
)
".ToCodeString();
            var expected = @"foo = | _
    (2 + 2) + 42".ToCodeString();

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
        public void WhenBackspacingClearsLineContinuatedCaretLine_PlacesCaretInsideStringOnPreviousLine()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = "foo = \"test\" & _\r\n     \"|\"".ToCodeString();
            var expected = "foo = \"test|\"".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void WhenBackspacingClearsLineContinuatedCaretLine_WithConcatenatedVbNewLine_PlacesCaretRightOfVbNewLineOnPreviousLine()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = "foo = \"test\" & vbNewLine & _\r\n     \"|\"".ToCodeString();
            var expected = "foo = \"test\" & vbNewLine|".ToCodeString();

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
            Assert.AreEqual(result ,default);
        }

        [Test]
        public void UnhandledKey_ReturnsDefault()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.A;
            var original = @"MsgBox |".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(result, default);
        }

        [Test]
        public void GivenClosingCharForUnmatchedOpeningChar_AsymetricPairBailsOut()
        {
            var pair = new SelfClosingPair('(', ')');
            if (pair.IsSymetric)
            {
                Assert.Inconclusive("Pair symetry is inconsistent with the purpose of the test.");
            }

            var input = pair.ClosingChar;
            var original = "MsgBox (|".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsNull(result);
        }

        [Test]
        public void GivenClosingCharForUnmatchedOpeningChar_SymetricPairBailsOut()
        {
            var pair = new SelfClosingPair('"', '"');
            if (!pair.IsSymetric)
            {
                Assert.Inconclusive("Pair symetry is inconsistent with the purpose of the test.");
            }

            var input = pair.ClosingChar;
            var original = "MsgBox \"|".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsNull(result);
        }

        [Test]
        public void GivenClosingCharForUnmatchedOpeningCharNonConsecutive_SymetricPairBailsOut()
        {
            var pair = new SelfClosingPair('"', '"');
            if (!pair.IsSymetric)
            {
                Assert.Inconclusive("Pair symetry is inconsistent with the purpose of the test.");
            }

            var input = pair.ClosingChar;
            var original = "MsgBox \"foo|".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsNull(result);
        }

        [Test]
        public void GivenOpeningCharInsideTerminatedStringLiteral_BailsOut()
        {
            var pair = new SelfClosingPair('(', ')');
            if (pair.IsSymetric)
            {
                Assert.Inconclusive("Pair symetry is inconsistent with the purpose of the test.");
            }

            var input = pair.OpeningChar;
            var original = "MsgBox \"foo|\"".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsNull(result);
        }

        [Test]
        public void GivenOpeningCharInsideNonTerminatedStringLiteral_BailsOut()
        {
            var pair = new SelfClosingPair('(', ')');
            if (pair.IsSymetric)
            {
                Assert.Inconclusive("Pair symetry is inconsistent with the purpose of the test.");
            }

            var input = pair.OpeningChar;
            var original = "MsgBox \"foo|".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsNull(result);
        }

        [Test]
        public void GivenClosingCharForUnmatchedOpeningChar_SingleChar_SymetricPairBailsOut()
        {
            var pair = new SelfClosingPair('"', '"');
            if (!pair.IsSymetric)
            {
                Assert.Inconclusive("Pair symetry is inconsistent with the purpose of the test.");
            }

            var input = pair.ClosingChar;
            var original = "\"|".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsNull(result);
        }
    }
}
