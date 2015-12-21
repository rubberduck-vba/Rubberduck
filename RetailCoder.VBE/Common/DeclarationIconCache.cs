using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Common
{
    public class DeclarationIconCache
    {
        private static readonly IDictionary<Tuple<DeclarationType, Accessibility>, BitmapImage> Images;

        static DeclarationIconCache()
        {
            var types = Enum.GetValues(typeof (DeclarationType)).Cast<DeclarationType>();
            var accessibilities = Enum.GetValues(typeof (Accessibility)).Cast<Accessibility>();

            Images = types.SelectMany(t => accessibilities.Select(a => Tuple.Create(t, a)))
                .ToDictionary(key => key, key => new BitmapImage(GetIconUri(key.Item1, key.Item2)));
        }

        public BitmapImage this[Declaration declaration]
        {
            get
            {
                var key = Tuple.Create(declaration.DeclarationType, declaration.Accessibility);
                return Images[key];
            }
        }

        public static BitmapImage ComponentIcon(vbext_ComponentType componentType)
        {
            Tuple<DeclarationType, Accessibility> key;
            switch (componentType)
            {
                case vbext_ComponentType.vbext_ct_StdModule:
                    key = Tuple.Create(DeclarationType.Module, Accessibility.Public);
                    break;
                case vbext_ComponentType.vbext_ct_ClassModule:
                    key = Tuple.Create(DeclarationType.Class, Accessibility.Public);
                    break;
                case vbext_ComponentType.vbext_ct_Document:
                    key = Tuple.Create(DeclarationType.Document, Accessibility.Public);
                    break;
                case vbext_ComponentType.vbext_ct_MSForm:
                    key = Tuple.Create(DeclarationType.UserForm, Accessibility.Public);
                    break;
                default:
                    key = Tuple.Create(DeclarationType.Project, Accessibility.Public);
                    break;
            }

            return Images[key];
        }

        private static Uri GetIconUri(DeclarationType declarationType, Accessibility accessibility)
        {
            const string baseUri = @"../../Resources/Microsoft/PNG/";

            string path;
            switch (declarationType)
            {
                case DeclarationType.Module:
                    path = "VSObject_Module.png";
                    break;

                case DeclarationType.Document | DeclarationType.Class: 
                    path = "document.png";
                    break;
                
                case DeclarationType.UserForm | DeclarationType.Class | DeclarationType.Control:
                    path = "VSProject_Form.png";
                    break;

                case DeclarationType.Class | DeclarationType.Module:
                    path = "VSProject_Class.png";
                    break;

                case DeclarationType.Procedure | DeclarationType.Member:
                case DeclarationType.Function | DeclarationType.Member:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "VSObject_Method_Private.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "VSObject_Method_Friend.png";
                        break;
                    }

                    path = "VSObject_Method.png";
                    break;

                case DeclarationType.PropertyGet | DeclarationType.Property | DeclarationType.Function:
                case DeclarationType.PropertyLet | DeclarationType.Property | DeclarationType.Procedure:
                case DeclarationType.PropertySet | DeclarationType.Property | DeclarationType.Procedure:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "VSObject_Properties_Private.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "VSObject_Properties_Friend.png";
                        break;
                    }

                    path = "VSObject_Properties.png";
                    break;

                case DeclarationType.Parameter:
                    path = "VSObject_Field_Shortcut.png";
                    break;

                case DeclarationType.Variable:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "VSObject_Field_Private.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "VSObject_Field_Friend.png";
                        break;
                    }

                    path = "VSObject_Field.png";
                    break;

                case DeclarationType.Constant:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "VSObject_Constant_Private.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "VSObject_Constant_Friend.png";
                        break;
                    }

                    path = "VSObject_Constant.png";
                    break;

                case DeclarationType.Enumeration:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "VSObject_Enum_Private.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "VSObject_Enum_Friend.png";
                        break;
                    }

                    path = "VSObject_Enum.png";
                    break;

                case DeclarationType.EnumerationMember | DeclarationType.Constant:
                    path = "VSObject_EnumItem.png";
                    break;

                case DeclarationType.Event:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "VSObject_Event_Private.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "VSObject_Event_Friend.png";
                        break;
                    }

                    path = "VSObject_Event.png";
                    break;

                case DeclarationType.UserDefinedType:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "VSObject_ValueTypePrivate.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "VSObject_ValueType_Friend.png";
                        break;
                    }

                    path = "VSObject_ValueType.png";
                    break;

                case DeclarationType.UserDefinedTypeMember | DeclarationType.Variable:
                    path = "VSObject_Field.png";
                    break;

                case DeclarationType.LibraryProcedure | DeclarationType.Procedure:
                case DeclarationType.LibraryFunction | DeclarationType.Function:
                    path = "VSObject_Method_Shortcut.png";
                    break;

                case DeclarationType.LineLabel:
                    path = "VSObject_Constant_Shortcut.png";
                    break;

                case DeclarationType.Project:
                    path = "VSObject_Library.png";
                    break;

                default:
                    path = "VSObject_Structure.png";
                    break;
            }

            return new Uri(baseUri + path, UriKind.Relative);
        }

    }
}