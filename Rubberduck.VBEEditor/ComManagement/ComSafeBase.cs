using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

#if DEBUG
using System.Diagnostics;
using System.Runtime.InteropServices;
#endif

namespace Rubberduck.VBEditor.ComManagement
{
    public abstract class ComSafeBase : IComSafe
    {
#if DEBUG
        protected IEnumerable<string> Trace = null;
#endif

        public abstract void Add(ISafeComWrapper comWrapper);

        //We do not use GetHashCode because subclasses of SafeComWrapper<T> overwrite this method 
        //and we need to distinguish between individual instances.
        protected int GetComWrapperObjectHashCode(ISafeComWrapper comWrapper)
        {
            return RuntimeHelpers.GetHashCode(comWrapper);
        }

        public abstract bool TryRemove(ISafeComWrapper comWrapper);

        public void Dispose()
        {
            Dispose(true);
        }

        protected abstract void Dispose(bool disposing);

#if DEBUG
        /// <summary>
        /// Provide a serialized list of the COM Safe
        /// to make it easy to analyze what is inside
        /// the COM Safe at the different points of
        /// the session's lifetime.
        /// </summary>
        public void Serialize()
        {
            using (var stream = System.IO.File.AppendText($"comSafeOutput {DateTime.UtcNow:yyyyMMddhhmmss}.csv"))
            {
                stream.WriteLine(
                    "Ordinal\tKey\tCOM Wrapper Type\tWrapping Null?\tIUnknown Pointer Address\tLevel 1\tLevel 2\tLevel 3");
                var i = 0;
                foreach (var kvp in GetWrappers())
                {
                    var line = kvp.Value != null
                        ? $"{i++}\t{kvp.Key}\t\"{kvp.Value.GetType().FullName}\"\t\"{kvp.Value.IsWrappingNullReference}\"\t\"{(kvp.Value.IsWrappingNullReference ? "null" : GetPtrAddress(kvp.Value))}\"\t\"{string.Join("\"\t\"", Trace)}\""
                        : $"{i++}\t{kvp.Key}\t\"null\"\t\"null\"\t\"null\"\t\"{string.Join("\"\t\"", Trace)}\"";
                    stream.WriteLine(line);
                }
            }
        }

        protected abstract IDictionary<int, ISafeComWrapper> GetWrappers();

        protected static IEnumerable<string> GetStackTrace(int frames, int offset)
        {
            var list = new List<string>();
            var trace = new StackTrace();
            if ((trace.FrameCount - offset) < frames)
            {
                frames = (trace.FrameCount - offset);
            }

            for (var i = 1; i <= frames; i++)
            {
                var frame = trace.GetFrame(i + offset);
                var typeName = frame.GetMethod().DeclaringType?.FullName ?? string.Empty;
                var methodName = frame.GetMethod().Name;

                var qualifiedName = $"{typeName}{(typeName.Length > 0 ? "::" : string.Empty)}{methodName}";
                list.Add(qualifiedName);
            }

            return list;
        }

        protected static string GetPtrAddress(object target)
        {
            if (target == null)
            {
                return IntPtr.Zero.ToString();
            }

            if (!Marshal.IsComObject(target))
            {
                return "Not a COM object";
            }

            var pointer = IntPtr.Zero;
            try
            {
                pointer = Marshal.GetIUnknownForObject(target);
            }
            finally
            {
                if (pointer != IntPtr.Zero)
                {
                    Marshal.Release(pointer);
                }
            }

            return pointer.ToString();
        }
#endif
    }
}