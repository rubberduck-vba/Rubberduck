using System;
using Antlr4.Runtime;
using Rubberduck.Extensions;
using Rubberduck.Inspections;
using Rubberduck.Parsing;
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

        public NavigateCodeEventArgs(QualifiedModuleName qualifiedModuleName, Selection selection)
        {
            _qualifiedName = qualifiedModuleName;
            _selection = selection;
        }

        public    NavigateCodeEventArgs(QualifiedSelection qualifiedSelection)
            :this(qualifiedSelection.QualifiedName, qualifiedSelection.Selection)
        {
        }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }
    }
}