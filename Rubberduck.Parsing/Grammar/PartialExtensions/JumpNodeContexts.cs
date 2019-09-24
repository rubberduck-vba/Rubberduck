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
        public partial class GoToStmtContext : IJumpNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
            public IExtendedNode Target { get; set; }
        }

        public partial class OnErrorStmtContext : IJumpNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
            public IExtendedNode Target { get; set; }
        }

        public partial class ResumeStmtContext : IJumpNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
            public IExtendedNode Target { get; set; }
        }

        public partial class ReturnStmtContext : IJumpNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
            public IExtendedNode Target { get; set; }
        }

        public partial class GoSubStmtContext : IJumpNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
            public IExtendedNode Target { get; set; }
        }

        public partial class ExitStmtContext : IExitNode
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
    }
}
