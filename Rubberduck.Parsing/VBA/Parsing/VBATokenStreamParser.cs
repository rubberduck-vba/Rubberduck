using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public class VBATokenStreamParser : TokenStreamParserBase
    {
        public VBATokenStreamParser(IParsePassErrorListenerFactory sllErrorListenerFactory, IParsePassErrorListenerFactory llErrorListenerFactory) 
        :base(sllErrorListenerFactory, llErrorListenerFactory)
        {
        }

        protected override IParseTree Parse(ITokenStream tokenStream, PredictionMode predictionMode, IParserErrorListener errorListener)
        {
            var parser = new VBAParser(tokenStream);
            parser.Interpreter.PredictionMode = predictionMode;
            parser.AddErrorListener(errorListener);
            return parser.startRule();
        }
    }
}
