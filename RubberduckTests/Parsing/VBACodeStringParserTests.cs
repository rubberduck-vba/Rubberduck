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
            var parser = new VBACodeStringParser("test", inputCode);
            Assert.IsInstanceOf<IParseTree>(parser.ParseTree);
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
                var parser = new VBACodeStringParser("test", inputCode);
            });
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void CannotParse_CodeSnippet()
        {
            const string inputCode = @"MsgBox ""hi""";

            Assert.Throws<MainParseSyntaxErrorException>(() =>
            {
                var parser = new VBACodeStringParser("test", inputCode);
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
            var parser = new VBACodeStringParser("test", inputCode);
            var tree = parser.ParseTree;

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
            var parser = new VBACodeStringParser("test", inputCode);
            
            Assert.IsInstanceOf<TokenStreamRewriter>(parser.Rewriter);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void Parse_ExplicitSll()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var parser = new VBACodeStringParser("test", inputCode, VBACodeStringParser.ParserMode.Sll);

            Assert.IsInstanceOf<IParseTree>(parser.ParseTree);
        }

        [Test]
        [Category("VBACodeStringParser_Tests")]
        public void Parse_ExplicitLl()
        {
            const string inputCode = @"
Public Sub Foo
    MsgBox ""hi""
End Sub";
            var parser = new VBACodeStringParser("test", inputCode, VBACodeStringParser.ParserMode.Ll);

            Assert.IsInstanceOf<IParseTree>(parser.ParseTree);
        }
    }
}
