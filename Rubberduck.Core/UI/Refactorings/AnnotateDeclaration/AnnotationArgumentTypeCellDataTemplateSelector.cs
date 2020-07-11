using System.Windows;
using System.Windows.Controls;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    public class AnnotationArgumentTypeCellDataTemplateSelector : DataTemplateSelector
    {
        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (!(container is FrameworkElement element)
                || !(item is AnnotationArgumentViewModel typedArgument))
            {
                return null;
            }

            if (typedArgument.CanEditArgumentType)
            {
                return element.FindResource("ArgumentMultiTypeTemplate") as DataTemplate;
            }

            return element.FindResource("ArgumentSingleTypeTemplate") as DataTemplate;
        }
    }
}