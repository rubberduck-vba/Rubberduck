using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
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

        private static Uri GetIconUri(DeclarationType declarationType, Accessibility accessibility)
        {
            const string baseUri = @"../../Resources/Custom/PNG/";

            string path;
            switch (declarationType)
            {
                case DeclarationType.ProceduralModule:
                    path = "ObjectModule.png";
                    break;

                case DeclarationType.Document | DeclarationType.ClassModule: 
                    path = "Document.png";
                    break;
                
                case DeclarationType.UserForm | DeclarationType.ClassModule | DeclarationType.Control:
                    path = "ProjectForm.png";
                    break;

                case DeclarationType.ClassModule | DeclarationType.ProceduralModule:
                    path = "ObjectClass.png";
                    break;

                case DeclarationType.Procedure | DeclarationType.Member:
                case DeclarationType.Function | DeclarationType.Member:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "ObjectMethodPrivate.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "ObjectMethodFriend.png";
                        break;
                    }

                    path = "ObjectMethod.png";
                    break;

                case DeclarationType.PropertyGet | DeclarationType.Property | DeclarationType.Function:
                case DeclarationType.PropertyLet | DeclarationType.Property | DeclarationType.Procedure:
                case DeclarationType.PropertySet | DeclarationType.Property | DeclarationType.Procedure:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "ObjectPropertiesPrivate.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "ObjectPropertiesFriend.png";
                        break;
                    }

                    path = "ObjectProperties.png";
                    break;

                case DeclarationType.Parameter:
                    path = "ObjectFieldShortcut.png";
                    break;

                case DeclarationType.Variable:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "ObjectFieldPrivate.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "ObjectFieldFriend.png";
                        break;
                    }

                    path = "ObjectField.png";
                    break;

                case DeclarationType.Constant:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "ObjectConstantPrivate.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "ObjectConstantFriend.png";
                        break;
                    }

                    path = "ObjectConstant.png";
                    break;

                case DeclarationType.Enumeration:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "ObjectEnumPrivate.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "ObjectEnumFriend.png";
                        break;
                    }

                    path = "ObjectEnum.png";
                    break;

                case DeclarationType.EnumerationMember:
                    path = "ObjectEnumItem.png";
                    break;

                case DeclarationType.Event:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "ObjectEventPrivate.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "ObjectEventFriend.png";
                        break;
                    }

                    path = "ObjectEvent.png";
                    break;

                case DeclarationType.UserDefinedType:
                    if (accessibility == Accessibility.Private)
                    {
                        path = "ObjectValueTypePrivate.png";
                        break;
                    }
                    if (accessibility == Accessibility.Friend)
                    {
                        path = "ObjectValueTypeFriend.png";
                        break;
                    }

                    path = "ObjectValueType.png";
                    break;

                case DeclarationType.UserDefinedTypeMember:
                    path = "ObjectField.png";
                    break;

                case DeclarationType.LibraryProcedure | DeclarationType.Procedure:
                case DeclarationType.LibraryFunction | DeclarationType.Function:
                    path = "ObjectMethodShortcut.png";
                    break;

                case DeclarationType.LineLabel:
                    path = "ObjectConstantShortcut.png";
                    break;

                case DeclarationType.Project:
                    path = "ObjectLibrary.png";
                    break;

                default:
                    path = "ObjectStructure.png";
                    break;
            }

            return new Uri(baseUri + path, UriKind.Relative);
        }

    }
}
