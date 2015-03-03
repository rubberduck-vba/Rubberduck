using System;
using Antlr4.Runtime;
using Rubberduck.Inspections;
using Rubberduck.VBA;

namespace Rubberduck.UI
{
    public class NavigateCodeEventArgs : EventArgs
    {
        public NavigateCodeEventArgs(QualifiedModuleName qualifiedName, ParserRuleContext context)
        {
            _qualifiedName = qualifiedName;
            _selection = context.GetSelection();
        }

        public NavigateCodeEventArgs(QualifiedModuleName qualifiedModuleName, Extensions.Selection selection)
        {
            _qualifiedName = qualifiedModuleName;
            _selection = selection;
        }

        public    NavigateCodeEventArgs(Extensions.QualifiedSelection qualifiedSelection)
            :this(qualifiedSelection.QualifiedName, qualifiedSelection.Selection)
        {
        }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private readonly Extensions.Selection _selection;
        public Extensions.Selection Selection { get { return _selection; } }
    }
}