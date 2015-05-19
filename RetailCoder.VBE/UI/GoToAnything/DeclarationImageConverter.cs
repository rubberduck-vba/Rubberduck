using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.GoToAnything
{
    public class DeclarationImageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var type = (Declaration) value;
            var image = new BitmapImage(GetImageForDeclaration(type));
            // todo: transparency?
            return image;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        private Uri GetImageForDeclaration(Declaration declaration)
        {
            switch (declaration.DeclarationType)
            {
                case DeclarationType.Module:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Module.bmp");
                case DeclarationType.Class:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Class.bmp");
                case DeclarationType.Procedure:
                case DeclarationType.Function:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Method_Private.bmp");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Method_Friend.bmp");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Method.bmp");

                case DeclarationType.PropertyGet:
                case DeclarationType.PropertyLet:
                case DeclarationType.PropertySet:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Properties_Private.bmp");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Properties_Friend.bmp");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Properties.bmp");

                case DeclarationType.Parameter:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Field_Private.bmp");
                case DeclarationType.Variable:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Field_Private.bmp");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Field_Friend.bmp");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Field.bmp");

                case DeclarationType.Constant:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Constant_Private.bmp");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Constant_Friend.bmp");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Constant.bmp");

                case DeclarationType.Enumeration:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Enum_Private.bmp");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Enum_Friend.bmp");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Enum.bmp");

                case DeclarationType.EnumerationMember:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_EnumItem.bmp");

                case DeclarationType.Event:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Event_Private.bmp");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Event_Friend.bmp");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Event.bmp");

                case DeclarationType.UserDefinedType:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Type_Private.bmp");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Type_Friend.bmp");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Type.bmp");

                case DeclarationType.UserDefinedTypeMember:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Field.bmp");

                case DeclarationType.LibraryProcedure:
                case DeclarationType.LibraryFunction:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Method.bmp");

                default:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/VSObject_Structure.bmp");
            }
        }
    }
}