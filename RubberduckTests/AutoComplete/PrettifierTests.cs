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
        [Test][Category("AutoComplete")]
        public void GivenSamePrettifiedCode_YieldsSameCodeString()
        {
            var original = "MsgBox (|".ToCodeString();
            var module = new Mock<ICodeModule>();
            module.Setup(m => m.GetLines(original.SnippetPosition)).Returns(original.Code);

            var sut = new CodeStringPrettifier(module.Object);
            var actual = sut.Prettify(original);

            Assert.AreEqual(original, actual);
        }

        [Test][Category("AutoComplete")]
        public void GivenTrailingWhitespace_PrettifiedCaretIsAtLastCharacter()
        {
            var original = "MsgBox |".ToCodeString();
            var prettified = "MsgBox".ToCodeString();
            var expected = "MsgBox|".ToCodeString();
            var module = new Mock<ICodeModule>();
            module.Setup(m => m.GetLines(original.SnippetPosition)).Returns(prettified.Code);

            var sut = new CodeStringPrettifier(module.Object);
            var actual = sut.Prettify(original);

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("AutoComplete")]
        public void GivenExtraWhitespace_PrettifiedCaretStillAtSameToken()
        {
            var original = "MsgBox      (\"test|\")".ToCodeString();
            var prettified = "MsgBox (\"test\")".ToCodeString();
            var expected = "MsgBox (\"test|\")".ToCodeString();

            var module = new Mock<ICodeModule>();
            module.Setup(m => m.GetLines(original.SnippetPosition)).Returns(prettified.Code);

            var sut = new CodeStringPrettifier(module.Object);
            var actual = sut.Prettify(original);

            Assert.AreEqual(expected, actual);
        }
    }
}