using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.Extensions
{
    public static class ComponentTypeExtensions
    {
        //VBA and VB6
        public const string ClassExtension = ".cls";
        public const string FormExtension = ".frm";
        public const string StandardExtension = ".bas";

        //VBA only
        public const string DocClassExtension = ".doccls";

        //VB6 only
        public const string UserControlExtension = ".ctl";
        public const string PropertyPageExtension = ".pag";
        public const string DocObjectExtension = ".dob";
        public const string ActiveXDesignerExtension = ".dsr";
        public const string ResourceExtension = ".res";

        //Binary resources
        public const string FormBinaryExtension = ".frx";
        public const string UserControlBinaryExtension = ".ctx";
        public const string PropertyPageBinaryExtension = ".pgx";
        public const string DocObjectBinaryExtension = ".dox";
        public const string ActiveXDesignerBinaryExtension = ".dsx";

        /// <summary>
        /// Returns the proper file extension for the Component Type.
        /// </summary>
        /// <remarks>Document classes should properly have a ".cls" file extension.
        /// However, because they cannot be removed and imported like other component types, we need to make a distinction.</remarks>
        /// <param name="componentType"></param>
        /// <returns>File extension that includes a preceding "dot" (.) </returns>
        public static string FileExtension(this ComponentType componentType)
        {
            foreach (var (extension, componentTypes) in VBAComponentTypesForExtension.Concat(VB6ComponentTypesForExtension))
            {
                if (componentTypes.Contains(componentType))
                {
                    return extension;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Returns the proper file extension for the binary file for the Component Type.
        /// </summary>
        /// <remarks>Returns an empty string if there is no binary file for the Component Type.</remarks>
        /// <param name="componentType"></param>
        /// <returns>File extension that includes a preceding "dot" (.) </returns>
        public static string BinaryFileExtension(this ComponentType componentType)
        {
            if (BinaryResourceExtensionForComponentType.TryGetValue(componentType, out var extension))
            {
                return extension;
            }

            return string.Empty;
        }

        public static IDictionary<string, ICollection<ComponentType>> ComponentTypesForExtension(VBEKind vbeKind)
        {
            return vbeKind == VBEKind.Hosted
                ? VBAComponentTypesForExtension
                : VB6ComponentTypesForExtension;
        }

        private static readonly IDictionary<string, ICollection<ComponentType>> VBAComponentTypesForExtension = new Dictionary<string, ICollection<ComponentType>>
        {
            [StandardExtension] = new List<ComponentType>{ComponentType.StandardModule},
            [ClassExtension] = new List<ComponentType>{ComponentType.ClassModule},
            [FormExtension] = new List<ComponentType>{ComponentType.UserForm},
            [DocClassExtension] = new List<ComponentType>{ComponentType.Document}
        };

        private static readonly IDictionary<string, ICollection<ComponentType>> VB6ComponentTypesForExtension = new Dictionary<string, ICollection<ComponentType>>
        {
            [StandardExtension] = new List<ComponentType>{ComponentType.StandardModule},
            [ClassExtension] = new List<ComponentType>{ComponentType.ClassModule},
            [FormExtension] = new List<ComponentType>{ComponentType.VBForm, ComponentType.MDIForm},
            [UserControlExtension] = new List<ComponentType>{ComponentType.UserControl},
            [PropertyPageExtension] = new List<ComponentType>{ComponentType.PropPage},
            [DocObjectExtension] = new List<ComponentType>{ComponentType.DocObject},
        };

        private static readonly IDictionary<ComponentType, string> BinaryResourceExtensionForComponentType = new Dictionary<ComponentType, string>
        {
            [ComponentType.UserForm] = FormBinaryExtension,
            [ComponentType.VBForm] = FormBinaryExtension,
            [ComponentType.MDIForm] = FormBinaryExtension,
            [ComponentType.UserControl] = UserControlBinaryExtension,
            [ComponentType.PropPage] = PropertyPageBinaryExtension,
            [ComponentType.DocObject] = DocObjectBinaryExtension,
            [ComponentType.ActiveXDesigner] = ActiveXDesignerBinaryExtension
        };
    }
}
