using System;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public class NavigateCodeIssueEventArgs : EventArgs
    {
        public NavigateCodeIssueEventArgs(SyntaxTreeNode node)
        {
            _node = node;
        }

        private readonly SyntaxTreeNode _node;
        public SyntaxTreeNode Node { get { return _node; } }
    }
}