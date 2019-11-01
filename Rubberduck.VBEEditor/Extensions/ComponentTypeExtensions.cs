using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.Extensions
{
    public static class ComponentTypeExtensions
    {
        public const string ClassExtension = ".cls";
        public const string FormExtension = ".frm";
        public const string StandardExtension = ".bas";
        public const string FormBinaryExtension = ".frx";
        public const string DocClassExtension = ".doccls";

        //TODO: double check whether the guesses below are correct.
        public const string UserControlExtension = ".ctl";
        public const string PropertyPageExtension = ".pag";
        public const string DocObjectExtension = ".dob";

        /// <summary>
        /// Returns the proper file extension for the Component Type.
        /// </summary>
        /// <remarks>Document classes should properly have a ".cls" file extension.
        /// However, because they cannot be removed and imported like other component types, we need to make a distinction.</remarks>
        /// <param name="componentType"></param>
        /// <param name="vbeKind"></param>
        /// <returns>File extension that includes a preceeding "dot" (.) </returns>
        public static string FileExtension(this ComponentType componentType)
        {
            switch (componentType)
            {
                case ComponentType.ClassModule:
                    return ClassExtension;
                case ComponentType.UserForm:
                    return FormExtension;
                case ComponentType.StandardModule:
                    return StandardExtension;
                case ComponentType.Document:
                    // documents should technically be a ".cls", but we need to be able to tell them apart.
                    return DocClassExtension;
                case ComponentType.PropPage:
                    return PropertyPageExtension;
                case ComponentType.UserControl:
                    return UserControlExtension;
                case ComponentType.DocObject:
                    return DocObjectExtension;
                default:
                    return string.Empty;
            }
        }

        public static IDictionary<string, ComponentType> ComponentTypeForExtension(VBEKind vbeKind)
        {
            return vbeKind == VBEKind.Hosted
                ? VBAComponentTypeForExtension
                : VB6ComponentTypeForExtension;
        }

        private static readonly IDictionary<string, ComponentType> VBAComponentTypeForExtension = new Dictionary<string, ComponentType>
        {
            [StandardExtension] = ComponentType.StandardModule,
            [ClassExtension] = ComponentType.ClassModule,
            [FormExtension] = ComponentType.UserForm,
            [DocClassExtension] = ComponentType.Document
        };

        private static readonly IDictionary<string, ComponentType> VB6ComponentTypeForExtension = new Dictionary<string, ComponentType>
        {
            [StandardExtension] = ComponentType.StandardModule,
            [ClassExtension] = ComponentType.ClassModule,
            [FormExtension] = ComponentType.VBForm,
            [UserControlExtension] = ComponentType.UserControl,
            [PropertyPageExtension] = ComponentType.PropPage,
            [DocObjectExtension] = ComponentType.DocObject,
        };
    }
}
