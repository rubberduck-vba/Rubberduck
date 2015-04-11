using System;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;

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

        public NavigateCodeEventArgs(Declaration declaration)
        {
            if (declaration == null)
            {
                return;
            }

            _declaration = declaration;
            _qualifiedName = declaration.QualifiedName.QualifiedModuleName;
            _selection = declaration.Selection;
        }

        public    NavigateCodeEventArgs(QualifiedSelection qualifiedSelection)
            :this(qualifiedSelection.QualifiedName, qualifiedSelection.Selection)
        {
        }

        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }
    }
}