using System;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public interface IContextFormatter
    {
        /// <summary>
        /// Determines the formatting of the contextual selection caption when a codepane is active.
        /// </summary>
        string Format(ICodePane activeCodePane, Declaration declaration);
        /// <summary>
        /// Determines the formatting of the contextual selection caption when a codepane is not active.
        /// </summary>
        string Format(Declaration declaration, bool multipleControls);
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
            var contextSelectionText = FormatDeclaration(declaration);

            return $"{codePaneSelectionText} | {contextSelectionText}";
        }

        public string Format(Declaration declaration, bool multipleControls)
        {
            return declaration == null ? string.Empty : FormatDeclaration(declaration, multipleControls);
        }

        private string FormatDeclaration(Declaration declaration, bool multipleControls = false)
        {
            var moduleName = declaration.QualifiedName.QualifiedModuleName;
            var declarationType = RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType, Settings.Settings.Culture);

            var typeName = TypeName(declaration, multipleControls, declarationType);
            var formattedDeclaration = FormattedDeclaration(declaration, typeName, moduleName, declarationType);
            return formattedDeclaration.Trim();
        }

        private static string FormattedDeclaration(
            Declaration declaration, 
            string typeName,
            QualifiedModuleName moduleName, 
            string declarationType)
        {
            if (declaration.ParentDeclaration != null)
            {
                if (declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
                {
                    // locals, parameters
                    return $"{declaration.ParentDeclaration.QualifiedName}:{declaration.IdentifierName} {typeName}";
                }

                if (declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
                {
                    // fields
                    var withEvents = declaration.IsWithEvents ? "(WithEvents) " : string.Empty;
                    return $"{withEvents}{moduleName}.{declaration.IdentifierName} {typeName}";
                }
            } 

            if (declaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                var formattedDeclaration = declaration.QualifiedName.ToString();
                if (declaration.DeclarationType == DeclarationType.Function
                    || declaration.DeclarationType == DeclarationType.PropertyGet)
                {
                    formattedDeclaration += typeName;
                }

                return formattedDeclaration;
            }
            
            if (declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                return $"{moduleName} ({declarationType})";
            }
            
            switch (declaration.DeclarationType)
            {
                case DeclarationType.Project:
                case DeclarationType.BracketedExpression:
                    var filename = System.IO.Path.GetFileName(declaration.QualifiedName.QualifiedModuleName.ProjectPath);
                    return $"{filename}{(string.IsNullOrEmpty(filename) ? string.Empty : ";")}{declaration.IdentifierName} ({declarationType})";
                case DeclarationType.Enumeration:
                case DeclarationType.UserDefinedType:
                    return !declaration.IsUserDefined
                        // built-in enums & UDT's don't have a module
                        ? $"{System.IO.Path.GetFileName(moduleName.ProjectPath)};{moduleName.ProjectName}.{declaration.IdentifierName}"
                        : moduleName.ToString();
                case DeclarationType.EnumerationMember:
                case DeclarationType.UserDefinedTypeMember:
                    return declaration.IsUserDefined
                        ? $"{moduleName}.{declaration.ParentDeclaration.IdentifierName}.{declaration.IdentifierName} {typeName}"
                        : $"{System.IO.Path.GetFileName(moduleName.ProjectPath)};{moduleName.ProjectName}.{declaration.ParentDeclaration.IdentifierName}.{declaration.IdentifierName} {typeName}";
                case DeclarationType.ComAlias:
                    return $"{System.IO.Path.GetFileName(moduleName.ProjectPath)};{moduleName.ProjectName}.{declaration.IdentifierName} (alias:{declaration.AsTypeName})";
            }

            return string.Empty;
        }

        private static string TypeName(Declaration declaration, bool multipleControls, string declarationType)
        {
            if (multipleControls)
            {
                return RubberduckUI.ContextMultipleControlsSelection;
            }

            var typeName = declaration.IsArray
                ? $"{declaration.AsTypeName}()"
                : declaration.AsTypeName;

            switch (declaration)
            {
                case ValuedDeclaration valued:
                    return $"({declarationType}{(string.IsNullOrEmpty(typeName) ? string.Empty : ":" + typeName)}{(string.IsNullOrEmpty(valued.Expression) ? string.Empty : $" = {valued.Expression}")})";
                case ParameterDeclaration parameter:
                    return $"({declarationType}{(string.IsNullOrEmpty(typeName) ? string.Empty : ":" + typeName)}{(string.IsNullOrEmpty(parameter.DefaultValue) ? string.Empty : $" = {parameter.DefaultValue}")})";
                default:
                    return $"({declarationType}{(string.IsNullOrEmpty(typeName) ? string.Empty : ":" + typeName)})";
            }
        }
    }
}