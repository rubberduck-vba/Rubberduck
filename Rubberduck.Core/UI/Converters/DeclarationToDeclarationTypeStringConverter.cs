using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UI.Converters
{
    public class DeclarationToDeclarationTypeStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var declaration = value as Declaration ?? (value as ICodeExplorerNode)?.Declaration;
            if (declaration == null)
            {
                return null;
            }

            if (value is CodeExplorerCustomFolderViewModel folder)
            {
                return folder.Description;
            }

            if (declaration is ClassModuleDeclaration classModule)
            {
                var supertype = classModule.SupertypeNames.FirstOrDefault();
                if (classModule.QualifiedModuleName.ComponentType ==
                    VBEditor.SafeComWrappers.ComponentType.Document)
                {
                    return $"{RubberduckUI.ResourceManager.GetString($"DeclarationType_{nameof(DeclarationType.Document)}", CultureInfo.CurrentUICulture)} ({classModule.IdentifierName}:{supertype})";
                }

                if (classModule.QualifiedModuleName.ComponentType == VBEditor.SafeComWrappers.ComponentType.UserForm)
                {
                    // a form has a predeclared ID, but we want it to have a UserForm icon:
                    return $"{RubberduckUI.ResourceManager.GetString($"DeclarationType_{nameof(DeclarationType.UserForm)}", CultureInfo.CurrentUICulture)} ({classModule.IdentifierName}:{supertype})";
                }

                if (classModule.IsInterface || classModule.Annotations.Any(annotation => annotation.Annotation is InterfaceAnnotation))
                {
                    return $"{RubberduckUI.ResourceManager.GetString($"DeclarationType_{nameof(DeclarationType.ClassModule)}", CultureInfo.CurrentUICulture)} (interface)";
                }

                if (classModule.HasPredeclaredId)
                {
                    return $"{RubberduckUI.ResourceManager.GetString($"DeclarationType_{nameof(DeclarationType.ClassModule)}", CultureInfo.CurrentUICulture)} (predeclared)";
                }

                return $"{RubberduckUI.ResourceManager.GetString($"DeclarationType_{nameof(DeclarationType.ClassModule)}", CultureInfo.CurrentUICulture)}";
            }

            if (declaration.DeclarationType.HasFlag(DeclarationType.ProceduralModule) && declaration.Annotations.Any(a => a.Annotation is TestModuleAnnotation))
            {
                return TestExplorer.TestExplorer_AddTestModule;
            }

            if (declaration.DeclarationType.HasFlag(DeclarationType.Member) && declaration.Annotations.Any(a => a.Annotation is TestMethodAnnotation))
            {
                return TestExplorer.TestExplorer_AddTestMethod;
            }

            return $"{RubberduckUI.ResourceManager.GetString($"DeclarationType_{declaration.DeclarationType}", CultureInfo.CurrentUICulture)}";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}