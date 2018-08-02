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

        [Test]
        public void SnippetPositionIsL1C1ifUnspecified()
        {
            var sut = new CodeString("|", new Selection());
            Assert.AreEqual(Selection.Home, sut.SnippetPosition);
        }

        [Test]
        public void NullCodeStringArgThrows()
        {
            Assert.Throws<ArgumentNullException>(() => new CodeString(null, Selection.Empty));
        }
    }
}
