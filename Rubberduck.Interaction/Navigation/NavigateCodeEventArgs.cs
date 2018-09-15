using System;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Interaction.Navigation
{
    public static class SelectionExtensions
    {
        public static NavigateCodeEventArgs GetNavitationArgs(this QualifiedSelection selection)
        {
            try
            {
                return new NavigateCodeEventArgs(new QualifiedSelection(selection.QualifiedName, selection.Selection));
            }
            catch (COMException)
            {
                return null;
            }
        }
    }

    public class NavigateCodeEventArgs : EventArgs
    {
        public NavigateCodeEventArgs(QualifiedModuleName qualifiedName, ParserRuleContext context)
        {
            QualifiedName = qualifiedName;
            Selection = context.GetSelection();
        }

        public NavigateCodeEventArgs(QualifiedModuleName qualifiedModuleName, Selection selection)
        {
            QualifiedName = qualifiedModuleName;
            Selection = selection;
        }

        public NavigateCodeEventArgs(Declaration declaration)
        {
            if (declaration == null)
            {
                return;
            }

            Declaration = declaration;
            QualifiedName = declaration.QualifiedName.QualifiedModuleName;
            Selection = declaration.Selection;
        }

        public NavigateCodeEventArgs(IdentifierReference reference)
        {
            if (reference == null)
            {
                return;
            }

            Reference = reference;
            QualifiedName = reference.QualifiedModuleName;
            Selection = reference.Selection;
        }

        public NavigateCodeEventArgs(QualifiedSelection qualifiedSelection)
            :this(qualifiedSelection.QualifiedName, qualifiedSelection.Selection)
        {
        }

        public IdentifierReference Reference { get; }

        public Declaration Declaration { get; }

        public QualifiedModuleName QualifiedName { get; }

        public Selection Selection { get; }
    }
}
