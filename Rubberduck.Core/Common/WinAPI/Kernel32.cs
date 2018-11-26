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
        public static extern ushort GlobalDeleteAtom(ushort nAtom);
        
        /// <summary>
        /// Sets the last-error code for the calling thread.
        /// </summary>
        /// <param name="dwErrorCode">The last-error code for the thread.</param>
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern void SetLastError(uint dwErrorCode);

        public static uint ERROR_SUCCESS = 0;
    }
}
