using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.FindSymbol
{
    public class DeclarationImageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var type = (Declaration) value;
            var image = new BitmapImage(GetImageForDeclaration(type));
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
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Module.png");
                case DeclarationType.Class:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Class.png");
                case DeclarationType.Procedure:
                case DeclarationType.Function:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Method_Private.png");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Method_Friend.png");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Method.png");

                case DeclarationType.PropertyGet:
                case DeclarationType.PropertyLet:
                case DeclarationType.PropertySet:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Properties_Private.png");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Properties_Friend.png");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Properties.png");

                case DeclarationType.Parameter:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Field_Private.png");
                case DeclarationType.Variable:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Field_Private.png");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Field_Friend.png");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Field.png");

                case DeclarationType.Constant:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Constant_Private.png");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Constant_Friend.png");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Constant.png");

                case DeclarationType.Enumeration:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Enum_Private.png");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Enum_Friend.png");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Enum.png");

                case DeclarationType.EnumerationMember:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_EnumItem.png");

                case DeclarationType.Event:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Event_Private.png");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Event_Friend.png");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Event.png");

                case DeclarationType.UserDefinedType:
                    if (declaration.Accessibility == Accessibility.Private)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Type_Private.png");
                    }
                    if (declaration.Accessibility == Accessibility.Friend)
                    {
                        return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Type_Friend.png");
                    }
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Type.png");

                case DeclarationType.UserDefinedTypeMember:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Field.png");

                case DeclarationType.LibraryProcedure:
                case DeclarationType.LibraryFunction:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Method.png");

                default:
                    return new Uri(@"pack://application:,,,/Rubberduck;component/Resources/Microsoft/PNG/VSObject_Structure.png");
            }
        }
    }
}