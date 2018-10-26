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
        public void PlacesCaretBetweenOpeningAndClosingChars_NestedPair()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = pair.OpeningChar;
            var original = "MsgBox (|)".ToCodeString();
            var expected = "MsgBox (\"|\")".ToCodeString();

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
        public void WhenNextPositionIsClosingChar_Nested_ClosingCharMovesSelection()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = pair.ClosingChar;
            var original = @"foo = MsgBox(""""|)".ToCodeString();
            var expected = @"foo = MsgBox("""")|".ToCodeString();

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
            var original = "Call xy(|z)".ToCodeString();
            var expected = "Call xy|z".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeletingOpeningChar_IndexExpr()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.Back;
            var original = "foo = CInt(|z)".ToCodeString();
            var expected = "foo = CInt|z".ToCodeString();

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

        [Test][Ignore("todo: figure out how to make this pass without breaking something else.")]
        public void BackspacingWorksWhenCaretIsNotOnLastNonEmptyLine_ConcatOnSameLine()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = "foo = \"\" & _\r\n      \"\" & _\r\n      \"|\" & _\r\n      \"\"".ToCodeString();
            var expected = "foo = \"\" & _\r\n      \"|\" & _\r\n      \"\"".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        [Ignore("todo: figure out how to make this pass without breaking something else.")]
        public void BackspacingWorksWhenCaretIsNotOnLastNonEmptyLine_ConcatOnNextLine()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = "foo = \"\" _\r\n      & \"\" _\r\n      & \"|\" _\r\n      & \"\"".ToCodeString();
            var expected = "foo = \"\" _\r\n      & \"|\" _\r\n      & \"\"".ToCodeString();

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
        public void WhenBackspacingClearsLineContinuatedCaretLine_WithConcatenatedVbNewLine_PlacesCaretInsideStringOnPreviousLine()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = "foo = \"test\" & vbNewLine & _\r\n     \"|\"".ToCodeString();
            var expected = "foo = \"test|\"".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void WhenBackspacingClearsLineContinuatedCaretLine_WithConcatenatedVbCrLf_PlacesCaretInsideStringOnPreviousLine()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = "foo = \"test\" & vbCrLf & _\r\n     \"|\"".ToCodeString();
            var expected = "foo = \"test|\"".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void WhenBackspacingClearsLineContinuatedCaretLine_WithConcatenatedVbCr_PlacesCaretInsideStringOnPreviousLine()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = "foo = \"test\" & vbCr & _\r\n     \"|\"".ToCodeString();
            var expected = "foo = \"test|\"".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void WhenBackspacingClearsLineContinuatedCaretLine_WithConcatenatedVbLf_PlacesCaretInsideStringOnPreviousLine()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = Keys.Back;
            var original = "foo = \"test\" & vbLf & _\r\n     \"|\"".ToCodeString();
            var expected = "foo = \"test|\"".ToCodeString();

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
        public void GivenEmptyIndentedLine_OpeningCharIsInsertedAtCaretPosition()
        {
            var pair = new SelfClosingPair('"', '"');
            var input = '"';
            var original = @"    |".ToCodeString();
            var expected = @"    ""|""".ToCodeString();

            var result = Run(pair, original, input);
            Assert.AreEqual(expected, result);
        }

        [Test]
        public void DeleteKey_ReturnsNull()
        {
            var pair = new SelfClosingPair('(', ')');
            var input = Keys.A;
            var original = @"MsgBox (|)".ToCodeString();

            var result = Run(pair, original, input);
            Assert.IsNull(result);
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
