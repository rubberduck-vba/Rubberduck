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
        public partial class EraseStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class MidStatementContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class LSetStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class RSetStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class RedimStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

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

        public partial class StopStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class WithStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class LineSpecialFormContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class CircleSpecialFormContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class ScaleSpecialFormContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class PSetSpecialFormContext: IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class CallStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class NameStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class ResetStmtContext : IExecutableNode
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

        public partial class SeekStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class LockStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class LineInputStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class WidthStmtContext : IExecutableNode
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

        public partial class PrintStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class WriteStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class InputStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class PutStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }

        public partial class GetStmtContext : IExecutableNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
        }
    }
}
