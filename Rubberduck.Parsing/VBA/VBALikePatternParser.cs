using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using NLog;
using Rubberduck.Parsing.Like;
using Rubberduck.Parsing.Symbols;
using System.Diagnostics;

namespace Rubberduck.Parsing.VBA
{
    public sealed class VBALikePatternParser
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Parses the given like pattern.
        /// </summary>
        /// <param name="likePattern">The like pattern of the like operation (e.g. in "a Like b" the b)</param>
        /// <returns>The root of the parse tree.</returns>
        public VBALikeParser.LikePatternStringContext Parse(string likePattern)
        {
            var stream = new AntlrInputStream(likePattern);
            var lexer = new VBALikeLexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBALikeParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            VBALikeParser.CompilationUnitContext tree = null;
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                tree = parser.compilationUnit();
            }
            catch
            {
                _logger.Warn("SLL mode failed for {0}. Retrying using LL.", likePattern);
                tokens.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.compilationUnit();
            }
            return tree.likePatternString();
        }
    }
}
