using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.Extensions
{
    public static class ComponentTypeExtensions
    {
        internal const string ClassExtension = ".cls";
        internal const string FormExtension = ".frm";
        internal const string StandardExtension = ".bas";
        internal const string FormBinaryExtension = ".frx";
        internal const string DocClassExtension = ".doccls";

        /// <summary>
        /// Returns the proper file extension for the Component Type.
        /// </summary>
        /// <remarks>Document classes should properly have a ".cls" file extension.
        /// However, because they cannot be removed and imported like other component types, we need to make a distinction.</remarks>
        /// <param name="componentType"></param>
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
                case ComponentType.ActiveXDesigner:
                default:
                    return string.Empty;
            }
        }
    }
}
