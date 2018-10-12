using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

#if DEBUG
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
#endif

namespace Rubberduck.VBEditor.ComManagement
{
    public class WeakComSafe : IComSafe
    {
        //We use weak references to allow the GC to reclaim RCWs earlier if possible.
        private readonly ConcurrentDictionary<int, (DateTime insertTime, WeakReference<ISafeComWrapper> weakRef)> _comWrapperCache = new ConcurrentDictionary<int, (DateTime, WeakReference<ISafeComWrapper>)>();

#if DEBUG
        private IEnumerable<string> trace = null;
#endif

        public void Add(ISafeComWrapper comWrapper)
        {
            if (comWrapper != null)
            {
#if DEBUG
                trace = GetStackTrace(3, 3);
#endif
                _comWrapperCache.AddOrUpdate(
                    GetComWrapperObjectHashCode(comWrapper), 
                    key => (DateTime.UtcNow, new WeakReference<ISafeComWrapper>(comWrapper)),
                    (key, value) =>
                    {
#if DEBUG
                        Debug.Assert(false);
#endif
                        return (value.insertTime, new WeakReference<ISafeComWrapper>(comWrapper));
                    });
            }

        }

        //We do not use GetHashCode because subclasses of SafeComWrapper<T> overwrite this method 
        //and we need to distinguish between individual instances.
        private int GetComWrapperObjectHashCode(ISafeComWrapper comWrapper)
        {
            return RuntimeHelpers.GetHashCode(comWrapper);
        }

        public bool TryRemove(ISafeComWrapper comWrapper)
        {
            return !_disposed && comWrapper != null && _comWrapperCache.TryRemove(GetComWrapperObjectHashCode(comWrapper), out _);
        }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            foreach (var weakReference in _comWrapperCache.Values)
            {
                if(weakReference.weakRef.TryGetTarget(out var comWrapper))
                {
                    comWrapper.Dispose();
                }
            }

            _comWrapperCache.Clear();
        }

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
                stream.WriteLine("Ordinal\tKey\tCOM Wrapper Type\tWrapping Null?\tIUnknown Pointer Address\tLevel 1\tLevel 2\tLevel 3");
                var i = 0;
                foreach (var kvp in _comWrapperCache.OrderBy(kvp => kvp.Value.insertTime))
                {
                    var line = kvp.Value.weakRef.TryGetTarget(out var target) 
                        ? $"{i++}\t{kvp.Key}\t\"{target.GetType().FullName}\"\t\"{target.IsWrappingNullReference}\"\t\"{(target.IsWrappingNullReference ? "null" : GetPtrAddress(target.Target))}\"\t\"{string.Join("\"\t\"", trace)}\"" 
                        : $"{i++}\t{kvp.Key}\t\"null\"\t\"null\"\t\"null\"\t\"{string.Join("\"\t\"", trace)}\"";
                    stream.WriteLine(line);
                }
            }
        }

        private static IEnumerable<string> GetStackTrace(int frames, int offset)
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

        private static string GetPtrAddress(object target)
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
