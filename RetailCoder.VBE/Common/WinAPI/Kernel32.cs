using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    /// <summary>
    /// Exposes Kernel32.dll API.
    /// </summary>
    public static class Kernel32
    {
        /// <summary>
        /// Adds a character string to the global atom table and returns a unique value (an atom) identifying the string.
        /// </summary>
        /// <param name="lpString">
        /// The null-terminated string to be added.
        /// The string can have a maximum size of 255 bytes.
        /// Strings that differ only in case are considered identical.
        /// The case of the first string of this name added to the table is preserved and returned by the GlobalGetAtomName function.
        /// </param>
        /// <returns>If the function succeeds, the return value is the newly created atom.</returns>
        [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)]
        public static extern ushort GlobalAddAtom(string lpString);

        /// <summary>
        /// Decrements the reference count of a global string atom. 
        /// If the atom's reference count reaches zero, GlobalDeleteAtom removes the string associated with the atom from the global atom table.
        /// </summary>
        /// <param name="nAtom">The atom and character string to be deleted.</param>
        /// <returns>The function always returns (ATOM) 0.</returns>
        [DllImport("kernel32.dll", SetLastError=true, ExactSpelling=true)]
        public static extern ushort GlobalDeleteAtom(IntPtr nAtom);

        /// <summary>
        /// Retrieves a module handle for the specified module. 
        /// The module must have been loaded by the calling process.
        /// </summary>
        /// <param name="lpModuleName">The name of the loaded module (either a .dll or .exe file). 
        /// If the file name extension is omitted, the default library extension .dll is appended. 
        /// The file name string can include a trailing point character (.) to indicate that the module name has no extension. 
        /// The string does not have to specify a path. When specifying a path, be sure to use backslashes (\), not forward slashes (/). 
        /// The name is compared (case independently) to the names of modules currently mapped into the address space of the calling process.</param>
        /// <returns>If the function succeeds, the return value is a handle to the specified module. 
        /// If the function fails, the return value is NULL. To get extended error information, call GetLastError.</returns>
        /// <remarks>The returned handle is not global or inheritable. It cannot be duplicated or used by another process.
        /// This function must be used carefully in a multithreaded application. There is no guarantee that the module handle remains valid between the time this function returns the handle and the time it is used. 
        /// For example, suppose that a thread retrieves a module handle, but before it uses the handle, a second thread frees the module. 
        /// If the system loads another module, it could reuse the module handle that was recently freed. 
        /// Therefore, the first thread would have a handle to a different module than the one intended.
        /// </remarks>
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);


    }
}
