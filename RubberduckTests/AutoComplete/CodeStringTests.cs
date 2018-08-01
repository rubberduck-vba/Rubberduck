using NUnit.Framework;
using Rubberduck.Common;
using Rubberduck.VBEditor;

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
}
