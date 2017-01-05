using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public interface IContextFormatter
    {
        /// <summary>
        /// Determines the formatting of the contextual selection caption.
        /// </summary>
        string Format(ICodePane activeCodePane, Declaration declaration);
    }

    public class ContextFormatter : IContextFormatter
    {
        public string Format(ICodePane activeCodePane, Declaration declaration)
        {
            if (activeCodePane == null)
            {
                return string.Empty;
            }

            var qualifiedSelection = activeCodePane.GetQualifiedSelection();
            if (declaration == null || !qualifiedSelection.HasValue)
            {
                return string.Empty;
            }

            var selection = qualifiedSelection.Value;
            var codePaneSelectionText = selection.Selection.ToString();
            var contextSelectionText = Format(declaration);

            return string.Format("{0} | {1}", codePaneSelectionText, contextSelectionText);
        }

        private string Format(Declaration declaration)
        {
            var formattedDeclaration = string.Empty;
            var moduleName = declaration.QualifiedName.QualifiedModuleName;
            var typeName = declaration.HasTypeHint
                ? Declaration.TypeHintToTypeName[declaration.TypeHint]
                : declaration.AsTypeName;
            var declarationType = RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType, Settings.Settings.Culture);

            typeName = "(" + declarationType + (string.IsNullOrEmpty(typeName) ? string.Empty : ":" + typeName) + ")";

            if (declaration.DeclarationType.HasFlag(DeclarationType.Project))
            {
                formattedDeclaration = System.IO.Path.GetFileName(declaration.QualifiedName.QualifiedModuleName.ProjectPath) + ";" + declaration.IdentifierName;
            }
            else if (declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                formattedDeclaration = moduleName.ToString();
            }
            
            if (declaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                formattedDeclaration = declaration.QualifiedName.ToString();
                if (declaration.DeclarationType == DeclarationType.Function
                    || declaration.DeclarationType == DeclarationType.PropertyGet)
                {
                    formattedDeclaration += typeName;
                }
            }
            
            if (declaration.DeclarationType == DeclarationType.Enumeration
                || declaration.DeclarationType == DeclarationType.UserDefinedType)
            {
                formattedDeclaration = declaration.IsBuiltIn
                    // built-in enums & UDT's don't have a module
                    ? System.IO.Path.GetFileName(moduleName.ProjectPath) + ";" + moduleName.ProjectName + "." + declaration.IdentifierName
                    : moduleName.ToString();
            }
            else if (declaration.DeclarationType == DeclarationType.EnumerationMember
                || declaration.DeclarationType == DeclarationType.UserDefinedTypeMember)
            {
                formattedDeclaration = string.Format("{0}.{1}.{2}",
                    declaration.IsBuiltIn 
                        ? System.IO.Path.GetFileName(moduleName.ProjectPath) + ";" + moduleName.ProjectName 
                        : moduleName.ToString(), 
                    declaration.ParentDeclaration.IdentifierName, 
                    declaration.IdentifierName);
            }

            var subscripts = declaration.IsArray ? "()" : string.Empty;
            if (declaration.ParentDeclaration != null && declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                // locals, parameters
                formattedDeclaration = string.Format("{0}:{1}{2} {3}", declaration.ParentDeclaration.QualifiedName, declaration.IdentifierName, subscripts, typeName);
            }

            if (declaration.ParentDeclaration != null && declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                // fields
                var withEvents = declaration.IsWithEvents ? "(WithEvents) " : string.Empty;
                formattedDeclaration = string.Format("{0}{1}.{2} {3}", withEvents, moduleName, declaration.IdentifierName, typeName);
            }

            return string.Format("{0} | {1} {2}", formattedDeclaration.Trim(), declaration.References.Count(), RubberduckUI.ContextReferences);
        }
    }
}