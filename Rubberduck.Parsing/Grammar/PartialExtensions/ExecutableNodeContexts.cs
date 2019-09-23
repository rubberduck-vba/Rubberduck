using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class CallStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class ExitStmtContext : IExecutableNode
        {
            public ExitStmtContext()
            {
                ExitsScope = (this.EXIT_SUB()
                        ?? this.EXIT_FUNCTION()
                        ?? this.EXIT_PROPERTY()
                        ) != null;
            }

            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool ExitsScope { get; }

            public bool IsReachable { get; set; }
        }

        public partial class RaiseEventStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class DebugPrintStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class OpenStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class CloseStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }
    }
}
