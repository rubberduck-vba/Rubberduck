using Moq;
using NUnit.Framework;
using Rubberduck.AutoComplete.SelfClosingPairCompletion;
using Rubberduck.Common;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.AutoComplete
{
    [TestFixture]
    public class PrettifierTests
    {
        [Test]
        [Category("AutoComplete")]
        public void ActuallyDeletesAndInsertsOriginalLine()
        {
            var code = "MsgBox|".ToCodeString();

            var sut = InitializeSut(code, code, out var module, out _);
            sut.Prettify(code);

            module.Verify(m => m.DeleteLines(code.SnippetPosition.StartLine, code.SnippetPosition.LineCount), Times.Once);
            module.Verify(m => m.InsertLines(code.SnippetPosition.StartLine, code.Code), Times.Once);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenSamePrettifiedCode_YieldsSameCodeString()
        {
            var original = "MsgBox (|".ToCodeString();

            var sut = InitializeSut(original, original, out _, out _);
            var actual = new TestCodeString(sut.Prettify(original));

            Assert.AreEqual(original, actual);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenLeadingWhitespace_YieldsSameCodeString()
        {
            var original = "    MsgBox|".ToCodeString();

            var sut = InitializeSut(original, original, out _, out _);
            var actual = new TestCodeString(sut.Prettify(original));

            Assert.AreEqual(original, actual);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenTrailingWhitespace_IsTrimmedAndPrettifiedCaretIsAtLastCharacter()
        {
            var original = "MsgBox |".ToCodeString();
            var prettified = "MsgBox".ToCodeString();
            var expected = "MsgBox|".ToCodeString();

            var sut = InitializeSut(original, prettified, out _, out _);
            var actual = new TestCodeString(sut.Prettify(original));

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenExtraWhitespace_PrettifiedCaretStillAtSameToken()
        {
            var original = "MsgBox      (\"test|\")".ToCodeString();
            var prettified = "MsgBox (\"test\")".ToCodeString();
            var expected = "MsgBox (\"test|\")".ToCodeString();

            var sut = InitializeSut(original, prettified, out _, out _);
            var actual = new TestCodeString(sut.Prettify(original));

            Assert.AreEqual(expected, actual);
        }

        private static CodeStringPrettifier InitializeSut(TestCodeString original, TestCodeString prettified, out Mock<ICodeModule> module, out Mock<ICodePane> pane)
        {
            module = new Mock<ICodeModule>();
            pane = new Mock<ICodePane>();
            pane.SetupProperty(m => m.Selection);
            module.Setup(m => m.DeleteLines(original.SnippetPosition.StartLine, original.SnippetPosition.LineCount));
            module.Setup(m => m.InsertLines(original.SnippetPosition.StartLine, original.Code));
            module.Setup(m => m.CodePane).Returns(pane.Object);
            module.Setup(m => m.GetLines(original.SnippetPosition)).Returns(prettified.Code);

            var sut = new CodeStringPrettifier(module.Object);
            return sut;
        }
    }
}