using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using Rubberduck.Parsing.Date;
using Rubberduck.Parsing.Symbols;
using System.Diagnostics;

namespace Rubberduck.Parsing.VBA
{
    public sealed class VBADateLiteralParser
    {
        /// <summary>
        /// Parses the given date.
        /// </summary>
        /// <param name="date">The date in string format including "hash tags" (e.g. #01-01-1900#)</param>
        /// <returns>The root of the parse tree.</returns>
        public VBADateParser.DateLiteralContext Parse(string date)
        {
            var stream = new AntlrInputStream(date);
            var lexer = new VBADateLexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBADateParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            VBADateParser.CompilationUnitContext tree = null;
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                tree = parser.compilationUnit();
            }
            catch
            {
                Debug.WriteLine(string.Format("{0}: SLL mode failed for {1}. Retrying using LL.", this.GetType().Name, date));
                tokens.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.compilationUnit();
            }
            return tree.dateLiteral();
        }
    }
}
