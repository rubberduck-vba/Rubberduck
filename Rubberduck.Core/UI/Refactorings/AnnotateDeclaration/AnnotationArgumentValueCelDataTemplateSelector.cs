using System.Windows;
using System.Windows.Controls;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.UI.Refactorings.AnnotateDeclaration
{
    public class AnnotationArgumentValueCellDataTemplateSelector : DataTemplateSelector
    {
        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (!(container is FrameworkElement element)
                || !(item is AnnotationArgumentViewModel typedArgument))
            {
                return null;
            }

            switch (typedArgument.ArgumentType)
            {
                case AnnotationArgumentType.Inspection:
                    return element.FindResource("ArgumentValueInspectionTemplate") as DataTemplate;
                case AnnotationArgumentType.Boolean:
                    return element.FindResource("ArgumentValueBooleanTemplate") as DataTemplate;
                default:
                    return element.FindResource("ArgumentValueTemplate") as DataTemplate;
            }
        }
    }
}