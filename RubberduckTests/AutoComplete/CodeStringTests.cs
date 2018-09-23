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
        public void ToStringIncludesCaretPipe()
        {
            var input = $"foo = MsgBox(|)";
            var sut = new TestCodeString(input, new Selection(0, input.IndexOf(TestCodeString.PseudoCaret)));

            Assert.AreEqual(input, sut.ToString());
        }

        [Test]
        public void CodeExcludesCaretPipe()
        {
            var input = $"foo = MsgBox(|)";
            var expected = "foo = MsgBox()";
            var sut = new TestCodeString(input, new Selection(0, input.IndexOf(TestCodeString.PseudoCaret)));

            Assert.AreEqual(expected, sut.Code);
        }

        [Test]
        public void SnippetPositionIsL1C1ifUnspecified()
        {
            var sut = new TestCodeString(TestCodeString.PseudoCaret.ToString(), new Selection());
            Assert.AreEqual(Selection.Home, sut.SnippetPosition);
        }

        [Test]
        public void NullCodeStringArgThrows()
        {
            Assert.Throws<ArgumentNullException>(() => new CodeString(null, Selection.Empty));
        }
    }
}
