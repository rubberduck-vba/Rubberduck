using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.Parsing.ComReflection
{
    public class ComLibraryProvider : IComLibraryProvider
    {
        #region Native Stuff
        // ReSharper disable InconsistentNaming
        // ReSharper disable UnusedMember.Local
        /// <summary>
        /// Controls how a type library is registered.
        /// </summary>
        private enum REGKIND
        {
            /// <summary>
            /// Use default register behavior.
            /// </summary>
            REGKIND_DEFAULT = 0,
            /// <summary>
            /// Register this type library.
            /// </summary>
            REGKIND_REGISTER = 1,
            /// <summary>
            /// Do not register this type library.
            /// </summary>
            REGKIND_NONE = 2
        }
        // ReSharper restore UnusedMember.Local

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
        private static extern int LoadTypeLibEx(string strTypeLibName, REGKIND regKind, out ITypeLib TypeLib);
        // ReSharper restore InconsistentNaming
        #endregion

        public ITypeLib LoadTypeLibrary(string libraryPath)
        {
            LoadTypeLibEx(libraryPath, REGKIND.REGKIND_NONE, out var typeLibrary);
            return typeLibrary;
        }
    }
}