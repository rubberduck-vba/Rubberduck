using Antlr4.Runtime;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class VBALexerTests
    {
        [TestCategory("Parser")]
        [TestMethod]
        public void TheLexerHidesLineNumbers()
        {
            string code = @"1 Sub Test()
2     For n = 1 To 10
3     Next n%
4
44 End Sub
23";
            string expectedTextOnDefaultChannel = @" Sub Test()
     For n = 1 To 10
     Next n%

 End Sub
";
            var tokenStream = TokenizedCode(code);
            tokenStream.Fill();
            var tokens = tokenStream.GetTokens();
            var defaultChannelTokens = tokens.Where(token => token.Channel == TokenConstants.DefaultChannel);
            var codeOnDefaultChannel = TokenText(defaultChannelTokens);

            Assert.AreEqual(expectedTextOnDefaultChannel, codeOnDefaultChannel);                        
        }

        [TestCategory("Parser")]
        [TestMethod]
        public void TheLexerDoesNotRemoveLineNumbers()
        {
            string code = @"1 Sub Test()
2     For n = 1 To 10
3     Next n%
4
44 End Sub
23";
            string expectedTextOnDefaultChannel = @"1 Sub Test()
2     For n = 1 To 10
3     Next n%
4
44 End Sub
23";
            var tokenStream = TokenizedCode(code);
            tokenStream.Fill();
            var tokens = tokenStream.GetTokens();
            var allCode = TokenText(tokens);

            Assert.AreEqual(expectedTextOnDefaultChannel, allCode);
        }

        private CommonTokenStream TokenizedCode(string code)
        {
            var stream = new AntlrInputStream(code);
            var lexer = new VBALexer(stream);
            return new CommonTokenStream(lexer);
        }

        private string TokenText(IEnumerable<IToken> tokens)
        {
            var builder = new StringBuilder();
            foreach (var token in tokens)
            {
                builder.Append(token.Text);
            }
            var withoutEOF = builder.ToString();
            while (withoutEOF.Length >= 5 && String.Equals(withoutEOF.Substring(withoutEOF.Length - 5, 5), "<EOF>"))
            {
                withoutEOF = withoutEOF.Substring(0, withoutEOF.Length - 5);
            }
            return withoutEOF;
        }
    }
}
