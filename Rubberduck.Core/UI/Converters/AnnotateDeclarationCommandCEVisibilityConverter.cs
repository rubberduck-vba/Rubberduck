using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Converters
{
    public class AnnotateDeclarationCommandCEVisibilityConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(values[0] is IAnnotation annotation)
                || !(values[1] is ICodeExplorerNode node))
            {
                return Visibility.Collapsed;
            }

            return ShouldBeVisible(annotation, node)
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        private bool ShouldBeVisible(IAnnotation annotation, ICodeExplorerNode node)
        {
            var target = node.Declaration;

            if (target == null)
            {
                return false;
            }

            if (!target.DeclarationType.HasFlag(DeclarationType.Module)
                && target.AttributesPassContext == null
                && annotation is IAttributeAnnotation)
            {
                return false;
            }

            var targetType = target.DeclarationType;

            switch (annotation.Target)
            {
                case AnnotationTarget.Member:
                    return targetType.HasFlag(DeclarationType.Member)
                           && targetType != DeclarationType.LibraryFunction
                           && targetType != DeclarationType.LibraryProcedure;
                case AnnotationTarget.Module:
                    return targetType.HasFlag(DeclarationType.Module);
                case AnnotationTarget.Variable:
                    return targetType.HasFlag(DeclarationType.Variable)
                           || targetType.HasFlag(DeclarationType.Constant);
                case AnnotationTarget.General:
                    return !targetType.HasFlag(DeclarationType.Module);
                case AnnotationTarget.Identifier:
                    return false;
                default:
                    return false;
            }
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}