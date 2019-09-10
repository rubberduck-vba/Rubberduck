using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar.Abstract;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class ArgContext : IIdentifierContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Identifier.GetName(this, out var tokenInterval);
                    return tokenInterval;
                }
            }
        }
    }
}
