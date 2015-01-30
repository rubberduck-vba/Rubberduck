using System;
using System.Runtime.InteropServices;
using Antlr4.Runtime;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public class NavigateCodeIssueEventArgs : EventArgs
    {
        public NavigateCodeIssueEventArgs(ParserRuleContext context)
        {
            _context = context;
        }

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }
    }
}