using System.Threading;
using Path = System.IO.Path;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.CodeAnalysis;
using System.Threading.Tasks;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public interface IContextFormatter
    {
        /// <summary>
        /// Determines the formatting of the contextual selection caption when a codepane is active.
        /// </summary>
        Task<string> FormatAsync(ICodePane activeCodePane, Declaration declaration, CancellationToken token);
        /// <summary>
        /// Determines the formatting of the contextual selection caption when a codepane is not active.
        /// </summary>
        Task<string> FormatAsync(Declaration declaration, bool multipleControls, CancellationToken token);
    }

    public class ContextFormatter : IContextFormatter
    {
        public async Task<string> FormatAsync(ICodePane activeCodePane, Declaration declaration, CancellationToken token)
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
            var contextSelectionText = await FormatDeclarationAsync(declaration, token);

            return $"{codePaneSelectionText} | {contextSelectionText}";
        }

        public async Task<string> FormatAsync(Declaration declaration, bool multipleControls, CancellationToken token)
        {
            if (declaration == null)
            {
                return string.Empty;
            }
            
            token.ThrowIfCancellationRequested();
            // designer, there is no code pane selection
            return await FormatDeclarationAsync(declaration, token, multipleControls);
        }

        private async Task<string> FormatDeclarationAsync(Declaration declaration, CancellationToken token, bool multipleControls = false)
        {
            token.ThrowIfCancellationRequested();
            var moduleName = declaration.QualifiedName.QualifiedModuleName;
            var declarationType = CodeAnalysisUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType, Settings.Settings.Culture);

            var typeName = TypeName(declaration, multipleControls, declarationType);
            var formattedDeclaration = await FormattedDeclarationAsync(declaration, typeName, moduleName, declarationType, token);
            return formattedDeclaration.Trim();
        }

        private async Task<string> FormattedDeclarationAsync(
            Declaration declaration, 
            string typeName,
            QualifiedModuleName moduleName, 
            string declarationType, 
            CancellationToken token)
        {
            return await Task.Run(() =>
            {

                token.ThrowIfCancellationRequested();
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
                        var withEvents = declaration.IsWithEvents ? $"({Tokens.WithEvents}) " : string.Empty;
                        return $"{withEvents}{moduleName}.{declaration.IdentifierName} {typeName}";
                    }
                }

                token.ThrowIfCancellationRequested();
                if (declaration.DeclarationType.HasFlag(DeclarationType.Member))
                {
                    var formattedDeclaration = $"{declaration.QualifiedName}";
                    if (declaration.DeclarationType == DeclarationType.Function
                        || declaration.DeclarationType == DeclarationType.PropertyGet)
                    {
                        formattedDeclaration += $" {typeName}";
                    }

                    return formattedDeclaration;
                }

                if (declaration.DeclarationType.HasFlag(DeclarationType.Module))
                {
                    return $"{moduleName} ({declarationType})";
                }

                token.ThrowIfCancellationRequested();
                switch (declaration.DeclarationType)
                {
                    case DeclarationType.Project:
                    case DeclarationType.BracketedExpression:
                        var filename = Path.GetFileName(declaration.QualifiedName.QualifiedModuleName.ProjectPath);
                        return
                            $"{filename}{(string.IsNullOrEmpty(filename) ? string.Empty : ";")}{declaration.IdentifierName} ({declarationType})";
                    case DeclarationType.Enumeration:
                    case DeclarationType.UserDefinedType:
                        return !declaration.IsUserDefined
                            // built-in enums & UDTs don't have a module
                            ? $"{Path.GetFileName(moduleName.ProjectPath)};{declaration.IdentifierName}"
                            : moduleName.ToString();
                    case DeclarationType.EnumerationMember:
                    case DeclarationType.UserDefinedTypeMember:
                        return declaration.IsUserDefined
                            ? $"{moduleName}.{declaration.ParentDeclaration.IdentifierName}.{declaration.IdentifierName} {typeName}"
                            : $"{Path.GetFileName(moduleName.ProjectPath)};{declaration.ParentDeclaration.IdentifierName}.{declaration.IdentifierName} {typeName}";
                    case DeclarationType.ComAlias:
                        return
                            $"{Path.GetFileName(moduleName.ProjectPath)};{declaration.IdentifierName} (alias:{declaration.AsTypeName})";
                }

                return string.Empty;
            }, token);
        }

        private static string TypeName(Declaration declaration, bool multipleControls, string declarationType)
        {
            if (multipleControls)
            {
                return RubberduckUI.ContextMultipleControlsSelection;
            }

            var typeName = Tokens.IDispatch.Equals(declaration.AsTypeName, System.StringComparison.InvariantCultureIgnoreCase)
                ? Tokens.Object
                : declaration.AsTypeName ?? string.Empty;

            var friendlyTypeName = declaration.IsArray ? $"{typeName}()" : typeName;

            switch (declaration)
            {
                case ValuedDeclaration valued:
                    return $"({declarationType}{(string.IsNullOrEmpty(friendlyTypeName) ? string.Empty : ":" + friendlyTypeName)}{(string.IsNullOrEmpty(valued.Expression) ? string.Empty : $" = {valued.Expression}")})";
                case ParameterDeclaration parameter:
                    return $"({declarationType}{(string.IsNullOrEmpty(friendlyTypeName) ? string.Empty : ":" + friendlyTypeName)}{(string.IsNullOrEmpty(parameter.DefaultValue) ? string.Empty : $" = {parameter.DefaultValue}")})";
                default:
                    return $"({declarationType}{(string.IsNullOrEmpty(friendlyTypeName) ? string.Empty : ":" + friendlyTypeName)})";
            }
        }
    }
}