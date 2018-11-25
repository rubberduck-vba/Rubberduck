using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    public class RubberduckParseErrorListenerBase : BaseErrorListener, IRubberduckParseErrorListener
    {
        public RubberduckParseErrorListenerBase(CodeKind codeKind)
        {
            CodeKind = codeKind;
        }

        protected CodeKind CodeKind { get; }
        
        //This serves as a method to postpone throwing a parse exception to after the entire input has been parsed,
        //e.g. when recovering from errors and collecting them.
        public virtual bool HasPostponedException(out Exception exception)
        {
            exception = null;
            return false;
        }
    }
}
