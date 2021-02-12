using Antlr4.Runtime;
using System;
using System.Text.RegularExpressions;

namespace Rubberduck.Parsing.Grammar
{
    // Currently this class does nothing, except allow other languages/implementations to define a custom contextSuperclass without having to change the grammar.
    public abstract class VBABaseParserRuleContext : ParserRuleContext
    {
        public VBABaseParserRuleContext() : base() { }
        
        public VBABaseParserRuleContext(ParserRuleContext parent, int invokingStateNumber) : base(parent, invokingStateNumber) { }
    }
}
