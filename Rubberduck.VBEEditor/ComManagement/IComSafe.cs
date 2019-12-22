using System;
using System.Diagnostics;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement
{
    public interface IComSafe: IDisposable
    {
        void Add(ISafeComWrapper comWrapper);
        bool TryRemove(ISafeComWrapper comWrapper);

        /// <summary>
        /// Available only if the compilation constant TRACE_COM_SAFE is set. Provide a mechanism for serializing both
        /// a snapshot of the COM safe at the instant and a historical activity log
        /// with a limited stack trace for each entry.
        /// </summary>
        /// <param name="targetDirectory">The path to a directory to place the serialized files in</param>
        void Serialize(string targetDirectory);
    }
}
