using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.VBEditor;
using System;

namespace RubberduckTests.AutoComplete
{
    [TestFixture]
    public class CodeStringTests
    {
        [Test]
        public void TestCodeString_ToStringIncludesCaretPipe()
        {
            var input = "foo = MsgBox(|)";
            var sut = input.ToCodeString();
            Assert.AreEqual(input, sut.ToString());
        }

        [Test]
        public void CodeExcludesCaretPipe()
        {
            var sut = "foo = MsgBox(|)".ToCodeString();
            var expected = "foo = MsgBox()";
            Assert.AreEqual(expected, sut.Code);
        }

        [Test]
        public void SnippetPositionIsL1C1IfUnspecified()
        {
            var sut = new TestCodeString(TestCodeString.PseudoCaret.ToString(), new Selection());
            Assert.AreEqual(Selection.Home, sut.SnippetPosition);
        }

        [Test]
        public void NullCodeStringArgThrows()
        {
            Assert.Throws<ArgumentNullException>(() => new CodeString(null, Selection.Empty));
        }

        [Test]
        public void IsInsideStringLiteral_TrueGivenCaretInsideSimpleString()
        {
            var sut = "foo = \"str|ing\"".ToCodeString();
            Assert.IsTrue(sut.IsInsideStringLiteral);
        }

        [Test]
        public void IsInsideStringLiteral_FalseGivenCaretOutsideSimpleString()
        {
            var sut = "foo = |\"string\"".ToCodeString();
            Assert.IsFalse(sut.IsInsideStringLiteral);
        }

        [Test]
        public void IsInsideStringLiteral_FalseGivenComment()
        {
            var sut = "'foo = \"|\"".ToCodeString();
            Assert.IsFalse(sut.IsInsideStringLiteral);
        }

        [Test]
        public void IsInsideStringLiteral_TrueGivenCaretInsideStringWithEscapedQuotes()
        {
            var sut = "foo = \"\"\"string|\"\"\"".ToCodeString();
            Assert.IsTrue(sut.IsInsideStringLiteral);
        }

        [Test]
        public void IsInsideStringLiteral_TrueGivenCaretInsideStringBetweenEscapedQuotes()
        {
            var sut = "foo = \"\"|\"string\"\"\"".ToCodeString();
            Assert.IsTrue(sut.IsInsideStringLiteral);
        }

        [Test]
        public void IsInsideStringLiteral_TrueGivenUnfinishedString()
        {
            var sut = "foo = \"unfinished string|".ToCodeString();
            Assert.IsTrue(sut.IsInsideStringLiteral);
        }

        [Test]
        public void IsComment_TrueGivenIsTrivialSingleQuoteComment()
        {
            var sut = "'\"not a string literal|, just a comment\"".ToCodeString();
            Assert.IsTrue(sut.IsComment);
        }

        [Test]
        public void IsComment_TrueGivenIsRemComment()
        {
            var sut = "Rem \"not a string literal|, just a comment\"".ToCodeString();
            Assert.IsTrue(sut.IsComment);
        }

        [Test]
        public void IsComment_TrueGivenIsRemCommentInSecondInstruction()
        {
            var sut = "foo = 2 + 2 : Rem \"not a string literal|, just a comment\"".ToCodeString();
            Assert.IsTrue(sut.IsComment);
        }

        [Test]
        public void IsComment_TrueGivenIsCommentStartingInPreviousPhysicalLine()
        {
            var sut = "' _\r\n\"not a string literal|, just a comment\"".ToCodeString();
            Assert.IsTrue(sut.IsComment);
        }

        [Test]
        public void IsComment_TrueGivenSingleQuoteCommentLine()
        {
            var sut = "'\"not a string literal|, just a comment\"".ToCodeString();
            Assert.IsTrue(sut.IsComment);
        }
    }
}
