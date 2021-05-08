using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.Converters
{
    public class CodeExplorerNodeToIconConverter : DeclarationToIconConverter, IMultiValueConverter
    {
        private static readonly ImageSource ProjectIcon = ToImageSource(CodeExplorerUI.ObjectLibrary);

        private static readonly ImageSource OpenFolderIcon = ToImageSource(CodeExplorerUI.FolderOpen);
        private static readonly ImageSource ClosedFolderIcon = ToImageSource(CodeExplorerUI.FolderClosed);

        private static readonly ImageSource ReferenceFolderIcon = ToImageSource(CodeExplorerUI.ObjectAssembly);
        private static readonly ImageSource ReferenceIcon = ToImageSource(CodeExplorerUI.Reference);

        private static readonly ImageSource LockedReferenceIcon = ToImageSource(CodeExplorerUI.LockedReference);
        private static readonly ImageSource BrokenReferenceIcon = ToImageSource(CodeExplorerUI.BrokenReference);


        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Declaration)
            {
                // invoked from code pane peek references command (no CE node)
                return base.Convert(value, targetType, parameter, culture);
            }

            if ((value as ICodeExplorerNode)?.Declaration is null)
            {
                return NullIconSource;
            }

            switch (value)
            {
                case CodeExplorerProjectViewModel _:
                    return ProjectIcon;
                case CodeExplorerReferenceFolderViewModel _:
                    return ReferenceFolderIcon;
                case CodeExplorerReferenceViewModel reference:
                    return reference.Reference is null || reference.Reference.IsBroken
                        ? BrokenReferenceIcon
                        : reference.Reference.IsBuiltIn
                            ? LockedReferenceIcon
                            : ReferenceIcon;
                case CodeExplorerCustomFolderViewModel folder:
                    return folder.IsExpanded ? OpenFolderIcon : ClosedFolderIcon;
                case CodeExplorerComponentViewModel component:
                    return base.Convert(component.Declaration, targetType, parameter, culture);
                default:
                    if (value is ICodeExplorerNode node)
                    {
                        return base.Convert(node.Declaration, targetType, parameter, culture);
                    }
                    return ExceptionIconSource;
            }
        }

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length == 0)
            {
                return null;
            };

            return Convert(values[0], targetType, parameter, culture);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
