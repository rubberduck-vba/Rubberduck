using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.Parsing.VBA;

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
            Assert.IsInstanceOf<IParseTree>(VBACodeStringParser.Parse(inputCode, e => e.startRule()).parseTree);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void CannotParse()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""";

            Assert.Throws<MainParseSyntaxErrorException>(() =>
            {
                VBACodeStringParser.Parse(inputCode, e => e.startRule());
            });
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void CannotParse_CodeSnippet()
        {
            const string inputCode = @"MsgBox ""hi""";

            Assert.Throws<MainParseSyntaxErrorException>(() =>
            {
                VBACodeStringParser.Parse(inputCode, e => e.startRule());
            });
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void ParseTreeIsValid()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var tree = VBACodeStringParser.Parse(inputCode, e => e.startRule());
            Assert.AreEqual(inputCode + "<EOF>", tree.parseTree.GetChild(0).GetText());
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void GetRewriter()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var tree = VBACodeStringParser.Parse(inputCode, e => e.startRule());

            Assert.IsInstanceOf<TokenStreamRewriter>(tree.rewriter);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void Parse_ExplicitSll()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var tree = VBACodeStringParser.Parse(inputCode, e => e.startRule());
            Assert.IsInstanceOf<IParseTree>(tree);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void Parse_ExplicitLl()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var tree = VBACodeStringParser.Parse(inputCode, e => e.startRule());
            Assert.IsInstanceOf<IParseTree>(tree);
        }
    }
}
