using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class SubStmtContext : IIdentifierContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval = Interval.Invalid;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
        }

        public partial class FunctionStmtContext : IIdentifierContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval = Interval.Invalid;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
        }

        public partial class EventStmtContext : IIdentifierContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval = Interval.Invalid;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
        }

        public partial class ArgContext : IIdentifierContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval = Interval.Invalid;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
        }

        public partial class VariableSubStmtContext : IIdentifierContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval = Interval.Invalid;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
        }

        public partial class PropertyGetStmtContext : IIdentifierContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval = Interval.Invalid;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
        }
        public partial class PropertyLetStmtContext : IIdentifierContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval = Interval.Invalid;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
        }
        public partial class PropertySetStmtContext : IIdentifierContext
        {
            public Interval IdentifierTokens
            {
                get
                {
                    Interval tokenInterval = Interval.Invalid;
                    Identifier.GetName(this, out tokenInterval);
                    return tokenInterval;
                }
            }
        }
    }
}
