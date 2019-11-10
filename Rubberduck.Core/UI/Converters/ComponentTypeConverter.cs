using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.Converters
{
    class ComponentTypeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ComponentType componentType)
            {
                switch (componentType)
                {
                    case ComponentType.ComComponent:
                        return Resources.RubberduckUI.ComponentType_ComComponent;
                    case ComponentType.Undefined:
                        return Resources.RubberduckUI.ComponentType_Undefined;
                    case ComponentType.StandardModule:
                        return Resources.RubberduckUI.ComponentType_StandardModule;
                    case ComponentType.ClassModule:
                        return Resources.RubberduckUI.ComponentType_ClassModule;
                    case ComponentType.UserForm:
                        return Resources.RubberduckUI.ComponentType_UserForm;
                    case ComponentType.ResFile:
                        return Resources.RubberduckUI.ComponentType_ResFile;
                    case ComponentType.VBForm:
                        return Resources.RubberduckUI.ComponentType_VBForm;
                    case ComponentType.MDIForm:
                        return Resources.RubberduckUI.ComponentType_MDIForm;
                    case ComponentType.PropPage:
                        return Resources.RubberduckUI.ComponentType_PropPage;
                    case ComponentType.UserControl:
                        return Resources.RubberduckUI.ComponentType_UserControl;
                    case ComponentType.DocObject:
                        return Resources.RubberduckUI.ComponentType_DocObject;
                    case ComponentType.RelatedDocument:
                        return Resources.RubberduckUI.ComponentType_RelatedDocument;
                    case ComponentType.ActiveXDesigner:
                        return Resources.RubberduckUI.ComponentType_ActiveXDesigner;
                    case ComponentType.Document:
                        return Resources.RubberduckUI.ComponentType_Document;
                    default:
                        return Binding.DoNothing;
                }
            }

            return Binding.DoNothing;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
