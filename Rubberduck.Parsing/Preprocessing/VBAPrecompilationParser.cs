using System;
using Antlr4.Runtime;
using Antlr4.Runtime.Atn;
using NLog;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.PreProcessing
{
    public sealed class VBAPrecompilationParser
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public VBAConditionalCompilationParser.CompilationUnitContext Parse(string moduleName, CommonTokenStream unprocessedTokenStream)
        {
            unprocessedTokenStream.Reset();
            var parser = new VBAConditionalCompilationParser(unprocessedTokenStream);
            parser.AddErrorListener(new ExceptionErrorListener()); // notify?
            VBAConditionalCompilationParser.CompilationUnitContext tree;
            try
            {
                parser.Interpreter.PredictionMode = PredictionMode.Sll;
                tree = parser.compilationUnit();
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "SLL mode failed in module {0}. Retrying using LL.", moduleName);
                unprocessedTokenStream.Reset();
                parser.Reset();
                parser.Interpreter.PredictionMode = PredictionMode.Ll;
                tree = parser.compilationUnit();
            }
            return tree;
        }
    }
}
