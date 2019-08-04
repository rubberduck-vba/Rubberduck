using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged
{
    /// <summary>
    /// Wraps the <see cref="Marshal" /> class so that we can perform
    /// additional actions such as logging or asserting. In production,
    /// it should be no different than just calling the class directly.
    /// </summary>
    /// <remarks>
    /// This is only a partial implementation - the <see cref="Marshal"/> class
    /// has many more methods; and if we need any, we should add it here and provide
    /// the appropriate logging and asserting for that method.
    /// 
    /// To avoid quote ambiguity, all strings returned by Debug.Print will
    /// be delimited with backtick (`).
    ///
    /// Due to performance and space considerations, the tracing requires use of additional flags
    /// TRACE_MARSHAL and REF_COUNT, which should be defined to enable the tracing.
    /// </remarks>
    internal static class RdMarshal
    {
        internal static int AddRef(IntPtr pUnk)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(pUnk!=IntPtr.Zero,"Null pointer passed in");

            Debug.Print($"Entering {nameof(AddRef)}; {nameof(pUnk)}: {FormatPtr(pUnk)}");
#endif
            var result = Marshal.AddRef(pUnk);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(AddRef)}; {nameof(result)}: {result}");
#endif
#if DEBUG && REF_COUNT
            PrintRefCount(pUnk);
#endif
            return result;
        }

        internal static IntPtr AllocHGlobal(int cb)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Entering {nameof(AllocHGlobal)}; {nameof(cb)}: {cb}");
#endif
            var result = Marshal.AllocHGlobal(cb);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(AllocHGlobal)}; {nameof(result)}: {result}");
#endif
#if DEBUG && REF_COUNT
            PrintAlloc(result);
#endif
            return result;
        }

        internal static void Copy(byte[] source, int startIndex, IntPtr destination, int length)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(destination != IntPtr.Zero, "Null pointer passed in");
            Debug.Assert(length >= 0, "Negative length passed in");
            Debug.Assert(startIndex >= 0, "Negative start index passed in");

            Debug.Print($"Executing {nameof(Copy)}; {nameof(source)}: {source.Length}, {nameof(startIndex)}: {startIndex}, {nameof(destination)}: {destination}, {nameof(length)}: {length}");
#endif
            Marshal.Copy(source, startIndex, destination, length);
        }

        internal static IntPtr CreateAggregatedObject<T>(IntPtr pOuter, T o)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(pOuter != IntPtr.Zero, "Null pointer passed in");

            Debug.Print($"Entering {nameof(CreateAggregatedObject)}; {nameof(pOuter)}: {pOuter}, {nameof(o)}: {o.GetType().Name}");
#endif
            var result = Marshal.CreateAggregatedObject(pOuter, o);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(CreateAggregatedObject)}; {nameof(result)}: {result}");
#endif
#if DEBUG && REF_COUNT
            PrintRefCount(result);
#endif
            return result;
        }

        internal static void FreeBSTR(IntPtr ptr)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(ptr!=IntPtr.Zero,"Null pointer passed in");
            Debug.Print($"Executing {nameof(FreeBSTR)}; {nameof(ptr)}: {FormatPtr(ptr)}");
#endif
#if DEBUG && REF_COUNT
            PrintFree(ptr);
#endif
            Marshal.FreeBSTR(ptr);
        }

        internal static void FreeHGlobal(IntPtr hglobal)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(hglobal != IntPtr.Zero, "Null pointer passed in");

            Debug.Print($"Executing {nameof(FreeHGlobal)}; {nameof(hglobal)}: {hglobal}");
#endif
#if DEBUG && REF_COUNT
            PrintFree(hglobal);
#endif
            Marshal.FreeHGlobal(hglobal);
        }

        internal static IntPtr GetComInterfaceForObject(object o, Type T)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Entering {nameof(GetComInterfaceForObject)}; {nameof(o)}: {o.GetType().Name}, {nameof(T)}: {T.Name}");
#endif
            var result = Marshal.GetComInterfaceForObject(o, T);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(GetComInterfaceForObject)}; {nameof(result)}: {result}");
#endif
#if DEBUG && REF_COUNT
            PrintRefCount(result);
#endif
            return result;
        }

        internal static IntPtr GetIUnknownForObject(object o)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Entering {nameof(GetIUnknownForObject)}; {nameof(o)}: {o.GetType().Name}");
#endif
            var result = Marshal.GetIUnknownForObject(o);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(GetIUnknownForObject)}; {nameof(result)}: {result}");
#endif
#if DEBUG && REF_COUNT
            PrintRefCount(result);
#endif
            return result;
        }

        internal static IntPtr GetIUnknownForObjectInContext(object o)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Entering {nameof(GetIUnknownForObjectInContext)}; {nameof(o)}: {o}");
#endif
            var result = Marshal.GetIUnknownForObjectInContext(o);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(GetIUnknownForObjectInContext)}; {nameof(result)}: {result}");
#endif
#if DEBUG && REF_COUNT
            PrintRefCount(result);
#endif
            return result;
        }

        internal static void GetNativeVariantForObject<T>(T obj, IntPtr pDstNativeVariant)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(pDstNativeVariant != IntPtr.Zero, "Null pointer passed in");

            Debug.Print($"Executing {nameof(GetNativeVariantForObject)}; {nameof(obj)}: {obj.GetType().Name}, {nameof(pDstNativeVariant)}: {pDstNativeVariant}");
#endif
            Marshal.GetNativeVariantForObject(obj, pDstNativeVariant);
        }

        internal static Exception GetExceptionForHR(int errorCode)
        {
            // Nothing interesting to write about but provided for parity with
            // the Marshal class.
            return Marshal.GetExceptionForHR(errorCode);
        }

        internal static object GetObjectForIUnknown(IntPtr pUnk)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(pUnk != IntPtr.Zero, "Null pointer passed in");

            Debug.Print($"Entering {nameof(GetObjectForIUnknown)}; {nameof(pUnk)}: {FormatPtr(pUnk)}");
#endif
            var result = Marshal.GetObjectForIUnknown(pUnk);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(GetObjectForIUnknown)}; {nameof(result)}: {result}");
#endif
            return result;
        }

        internal static object GetTypedObjectForIUnknown(IntPtr pUnk, Type t)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Executing {nameof(GetTypedObjectForIUnknown)}; {nameof(pUnk)}: {FormatPtr(pUnk)}, {nameof(t)}: `{t.Name}`");
#endif
            return Marshal.GetTypedObjectForIUnknown(pUnk, t);
        }

        internal static Type GetTypeForITypeInfo(IntPtr pTypeInfo)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(pTypeInfo != IntPtr.Zero, "Null pointer was passed in");

            Debug.Print($"Entering {nameof(GetTypeForITypeInfo)}; {nameof(pTypeInfo)}: {pTypeInfo}");
#endif
            var result = Marshal.GetTypeForITypeInfo(pTypeInfo);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(GetTypeForITypeInfo)}; {nameof(result)}: {result}");
#endif
            return result;
        }
        
        internal static string GetTypeLibName(ITypeLib typelib)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Entering {nameof(GetTypeLibName)};");
#endif
            var result = Marshal.GetTypeLibName(typelib);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(GetTypeLibName)}; {nameof(result)}: `{result}`");
#endif
            return result;
        }

        internal static bool IsComObject(object o)
        {
            // Nothing interesting to write about but provided for parity with
            // the Marshal class.
            return Marshal.IsComObject(o);
        }

        internal static string PtrToStringBSTR(IntPtr ptr)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(ptr!=IntPtr.Zero, "Null pointer passed in");

            Debug.Print($"Entering {nameof(PtrToStringBSTR)}; {nameof(ptr)}: {FormatPtr(ptr)}");
#endif
            var result = Marshal.PtrToStringBSTR(ptr);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(PtrToStringBSTR)}; {nameof(result)}: `{result}`");
#endif
            return result;
        }

        internal static object PtrToStructure(IntPtr ptr, Type T)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(ptr != IntPtr.Zero, "Null pointer passed in");

            Debug.Print($"Entering {nameof(PtrToStructure)}; {nameof(ptr)}: {FormatPtr(ptr)}, {nameof(T)}: {T.Name}");
#endif
            var result = Marshal.PtrToStructure(ptr, T);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(PtrToStructure)}");
#endif
            return result;
        }

        internal static int QueryInterface(IntPtr pUnk, ref Guid iid, out IntPtr ppv)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(pUnk != IntPtr.Zero, "Null pointer passed in");
            Debug.Assert(iid != Guid.Empty, "Empty IID passed in");

            Debug.Print($"Entering {nameof(QueryInterface)}; {nameof(pUnk)}: {FormatPtr(pUnk)}, {nameof(iid)}: {iid.ToString()}");
#endif
            var result = Marshal.QueryInterface(pUnk, ref iid, out ppv);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(QueryInterface)}; {nameof(result)}: {result}, {nameof(ppv)}: {ppv}");
#endif
#if DEBUG && REF_COUNT
            PrintRefCount(ppv);
#endif
            return result;
        }

        internal static int ReadInt32(IntPtr ptr)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(ptr!=IntPtr.Zero);

            Debug.Print($"Entering {nameof(ReadInt32)}; {nameof(ptr)}: {FormatPtr(ptr)}");
#endif
            var result = Marshal.ReadInt32(ptr);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(ReadInt32)}; {nameof(result)}: {result}");
#endif
            return result;
        }

        internal static IntPtr ReadIntPtr(IntPtr ptr)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(ptr != IntPtr.Zero);

            Debug.Print($"Entering {nameof(ReadIntPtr)}; {nameof(ptr)}: {FormatPtr(ptr)}");
#endif
            var result = Marshal.ReadIntPtr(ptr);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(ReadIntPtr)}; {nameof(result)}: {result}");
#endif
            return result;
        }

        internal static int Release(IntPtr pUnk)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Entering {nameof(Release)}; {nameof(pUnk)}: {FormatPtr(pUnk)}");
#endif
            var result = Marshal.Release(pUnk);
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(result >= 0, "The ref count is negative which is invalid.");
            Debug.Print($"Leaving {nameof(Release)}; {nameof(result)}: {result}");
#endif
#if DEBUG && REF_COUNT
            Debug.Print($"{nameof(Release)}:: COM Object: {FormatPtr(pUnk)} ref count {result}");
#endif
            return result;
        }

        internal static int ReleaseComObject(object o)
        {
#if DEBUG && TRACE_MARSHAL
            var ptr = Marshal.GetIUnknownForObject(o);
            Debug.Print($"Entering {nameof(ReleaseComObject)}; {nameof(o)}: {FormatPtr(ptr)}, {o.GetType().Name}");
            var debugResult = Marshal.Release(ptr);
#endif
            var result = Marshal.ReleaseComObject(o);
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(debugResult > 0,
                $"The ref count is at zero or is invalid before calling the {nameof(Marshal.ReleaseComObject)}.");
            Debug.Assert(result >= 0,
                $"The ref count is invalid after calling the {nameof(Marshal.ReleaseComObject)}");
            Debug.Print($"Leaving {nameof(ReleaseComObject)}; {nameof(result)}: {result}");
#endif
#if DEBUG && REF_COUNT
            Debug.Print($"{nameof(ReleaseComObject)}:: COM Object: {o.GetType().Name} ref count {result}");
#endif
            return result;
        }

        internal static int SizeOf(Type t)
        {
#if DEBUG && TRACE_MARSHAL  
            Debug.Print($"Entering {nameof(SizeOf)}; {nameof(t)}: {t.Name}");
#endif
            var result = Marshal.SizeOf(t);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(SizeOf)}; {nameof(result)}: {result}");
#endif
            return result;
        }

        internal static int SizeOf(object structure)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Entering {nameof(SizeOf)}; {nameof(structure)}: {structure.GetType().Name}");
#endif
            var result = Marshal.SizeOf(structure);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(SizeOf)}; {nameof(result)}: {result}");
#endif
            return result;
        }

        internal static IntPtr StringToBSTR(string s)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Entering {nameof(StringToBSTR)}; {nameof(s)}: `{s}`");
#endif
            var result = Marshal.StringToBSTR(s);
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Leaving {nameof(StringToBSTR)}; {nameof(result)}: {result}");
#endif
#if DEBUG && REF_COUNT
            PrintAlloc(result);
#endif
            return result;
        }

        internal static void StructureToPtr<T>(T structure, IntPtr ptr, bool fDeleteOld)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Print($"Executing {nameof(StructureToPtr)}; {nameof(structure)}: {structure.GetType().Name}, {nameof(ptr)}: {FormatPtr(ptr)}, {nameof(fDeleteOld)}: {fDeleteOld}");
#endif
            Marshal.StructureToPtr(structure, ptr, fDeleteOld);
#if DEBUG
            if (!fDeleteOld) { Debug.Print($"Warning: {nameof(StructureToPtr)} had flag {fDeleteOld} passed with false. That may cause memory leaks.");}
#endif
        }

        internal static void WriteInt32(IntPtr ptr, int val)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(ptr != IntPtr.Zero, "Null pointer passed in");

            Debug.Print($"Executing {nameof(WriteInt32)}; {nameof(ptr)}: {FormatPtr(ptr)}, {nameof(val)}: {val}");
#endif
            Marshal.WriteInt32(ptr, val);
        }

        internal static void WriteIntPtr(IntPtr ptr, IntPtr val)
        {
#if DEBUG && TRACE_MARSHAL
            Debug.Assert(ptr!= IntPtr.Zero, "Null pointer passed in");

            Debug.Print($"Executing {nameof(WriteIntPtr)}; {nameof(ptr)}: {FormatPtr(ptr)}, {nameof(val)}: {val}");
#endif
            Marshal.WriteIntPtr(ptr, val);
        }

        //Helper functions
        private static void PrintRefCount(IntPtr pUnk, [CallerMemberName] string methodName = null)
        {
            if (pUnk != IntPtr.Zero)
            {
                var refCount = Marshal.AddRef(pUnk) - 1;
                Marshal.Release(pUnk);
                Debug.Print($"{methodName}:: COM Object: {FormatPtr(pUnk)} ref count {refCount}");
            }
            else
            {
                Debug.Print($"{methodName}:: COM Object: not allocated");
            }
        }

        private static void PrintAlloc(IntPtr pUnmanaged, [CallerMemberName] string methodName = null)
        {
            Debug.Print($"{methodName}:: Unmanaged pointer allocated: {FormatPtr(pUnmanaged)}");
        }

        private static void PrintFree(IntPtr pUnmanaged, [CallerMemberName] string methodName = null)
        {
            Debug.Print($"{methodName}:: Unmanaged pointer released: {FormatPtr(pUnmanaged)}");
        }

        internal static string FormatPtr(IntPtr ptr)
        {
            return string.Concat("0x",
                (Marshal.SizeOf<IntPtr>() == 8 ? 
                    ptr.ToInt64().ToString("X16") : 
                    ptr.ToInt32().ToString("X8")));
        }
    }
}
