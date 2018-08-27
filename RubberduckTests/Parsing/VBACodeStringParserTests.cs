using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using NUnit.Framework;
using Rubberduck.Parsing.VBA.Parsing;

namespace RubberduckTests.Parsing
{
    [TestFixture]
    public class VBACodeStringParserTests
    {
        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void CanParse()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var result = VBACodeStringParser.Parse(inputCode, t => t.startRule());
            Assert.IsInstanceOf<IParseTree>(result.parseTree);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void ParseTreeIsValid()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var result = VBACodeStringParser.Parse(inputCode, t => t.startRule());
            var tree = result.parseTree;

            Assert.AreEqual(inputCode + "<EOF>", tree.GetChild(0).GetText());
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void GetRewriter()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var result = VBACodeStringParser.Parse(inputCode, t => t.startRule());
            Assert.IsInstanceOf<TokenStreamRewriter>(result.rewriter);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void Parse_ExplicitSll()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var result = VBACodeStringParser.Parse(inputCode, t => t.startRule(), VBACodeStringParser.ParserMode.Sll);
            Assert.IsInstanceOf<IParseTree>(result.parseTree);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void Parse_ExplicitLl()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var result = VBACodeStringParser.Parse(inputCode, t => t.startRule(), VBACodeStringParser.ParserMode.Ll);
            Assert.IsInstanceOf<IParseTree>(result.parseTree);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void CanParseEmptyModule()
        {
            const string inputCode = @"";

            var result = VBACodeStringParser.Parse(inputCode, t => t.startRule());
            Assert.IsInstanceOf<IParseTree>(result.parseTree);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void CanParseNullInput()
        {
            const string inputCode = null;

            var result = VBACodeStringParser.Parse(inputCode, t => t.startRule());
            Assert.IsInstanceOf<IParseTree>(result.parseTree);
        }
    }
}
