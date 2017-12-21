using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using NLog;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class VBAPrecompilationParser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public VBAConditionalCompilationParser.CompilationUnitContext Parse(string moduleName, PredictionMode predictionMode, CommonTokenStream unprocessedTokenStream, BaseErrorListener errorListener)
        {
            unprocessedTokenStream.Reset();
            var parser = new VBAConditionalCompilationParser(unprocessedTokenStream);
            parser.AddErrorListener(errorListener);
            parser.Interpreter.PredictionMode = predictionMode;
            parser.ErrorHandler = new RecoveryStrategy();
            var tree = parser.compilationUnit();
            return tree;
        }
    }
}
