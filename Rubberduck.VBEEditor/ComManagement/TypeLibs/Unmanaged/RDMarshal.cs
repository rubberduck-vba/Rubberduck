using System;
using System.Diagnostics;
using System.Linq;
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
            Trace(TraceAction.Entering, pUnk);
            var result = Marshal.AddRef(pUnk);
            Trace(TraceAction.Leaving, pUnk, args: (nameof(result), result));
            PrintRefCount(pUnk);
            return result;
        }
        
        internal static IntPtr AllocHGlobal(int cb)
        {
            Trace(TraceAction.Entering, args: (nameof(cb), cb));
            var result = Marshal.AllocHGlobal(cb);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            PrintAlloc(result);
            return result;
        }

        internal static void Copy(byte[] source, int startIndex, IntPtr destination, int length)
        {
            Assert(() =>
            {
                Debug.Assert(destination != IntPtr.Zero, "Null pointer passed in");
                Debug.Assert(length >= 0, "Negative length passed in");
                Debug.Assert(startIndex >= 0, "Negative start index passed in");
            });

            Trace(TraceAction.Executing, args: new (string, object)[]
            {
                (nameof(source), source.Length),
                (nameof(startIndex), startIndex),
                (nameof(destination), destination), 
                (nameof(length), length)
            });

            Marshal.Copy(source, startIndex, destination, length);
        }

        internal static IntPtr CreateAggregatedObject<T>(IntPtr pOuter, T o)
        {
            Assert(() =>
            {
                Debug.Assert(pOuter != IntPtr.Zero, "Null pointer passed in"); 
            });

            Trace(TraceAction.Entering,
                args: new (string Name, object Value)[]
                {
                    (nameof(pOuter), pOuter),
                    (nameof(o), o.GetType().Name)
                });

            var result = Marshal.CreateAggregatedObject(pOuter, o);

            Trace(TraceAction.Leaving, args: (nameof(result), result));
            PrintRefCount(result);

            return result;
        }

        internal static void FreeBSTR(IntPtr ptr)
        {

            Assert(() =>
            {
                Debug.Assert(ptr != IntPtr.Zero, "Null pointer passed in"); 
            });

            Trace(TraceAction.Executing, args: (nameof(ptr), FormatPtr(ptr)));
            PrintFree(ptr);

            Marshal.FreeBSTR(ptr);
        }

        internal static void FreeHGlobal(IntPtr hglobal)
        {
            Assert(() =>
            {
                Debug.Assert(hglobal != IntPtr.Zero, "Null pointer passed in");
            });

            Trace(TraceAction.Executing, args: (nameof(hglobal), FormatPtr(hglobal)));
            PrintFree(hglobal);

            Marshal.FreeHGlobal(hglobal);
        }

        internal static IntPtr GetComInterfaceForObject(object o, Type T)
        {
            Trace(TraceAction.Entering, args: new (string Name, object Value)[]
            {
                (nameof(o), o.GetType().Name), 
                (nameof(T), T.Name)
            });

            var result = Marshal.GetComInterfaceForObject(o, T);

            Trace(TraceAction.Leaving, args: (nameof(result), result));
            PrintRefCount(result);

            return result;
        }

        internal static IntPtr GetIUnknownForObject(object o)
        {
            Trace(TraceAction.Entering, args: (nameof(o), o.GetType().Name));
            var result = Marshal.GetIUnknownForObject(o);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            PrintRefCount(result);
            return result;
        }

        internal static IntPtr GetIUnknownForObjectInContext(object o)
        {
            Trace(TraceAction.Entering, args:(nameof(o), o));
            var result = Marshal.GetIUnknownForObjectInContext(o);
            Trace(TraceAction.Leaving, args:(nameof(result), result));
            PrintRefCount(result);
            return result;
        }

        internal static void GetNativeVariantForObject<T>(T obj, IntPtr pDstNativeVariant)
        {
            Assert(() =>
            {
                Debug.Assert(pDstNativeVariant != IntPtr.Zero, "Null pointer passed in"); 
            });

            Trace(TraceAction.Executing, args: new (string Name, object Value)[]
            {
                (nameof(obj), obj.GetType().Name), 
                (nameof(pDstNativeVariant), pDstNativeVariant)
            });
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
            Assert(() =>
            {
                Debug.Assert(pUnk != IntPtr.Zero, "Null pointer passed in");
            });

            Trace(TraceAction.Entering, args: (nameof(pUnk), FormatPtr(pUnk)));
            var result = Marshal.GetObjectForIUnknown(pUnk);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            return result;
        }

        internal static object GetTypedObjectForIUnknown(IntPtr pUnk, Type t)
        {
            Trace(TraceAction.Executing, args: new (string Name, object Value)[]
            {
                (nameof(pUnk), FormatPtr(pUnk)), 
                (nameof(t), $"`{t.Name}`")
            });

            return Marshal.GetTypedObjectForIUnknown(pUnk, t);
        }

        internal static Type GetTypeForITypeInfo(IntPtr pTypeInfo)
        {
            Assert(() =>
            {
                Debug.Assert(pTypeInfo != IntPtr.Zero, "Null pointer was passed in");
            });
            
            Trace(TraceAction.Entering, args: (nameof(pTypeInfo), FormatPtr(pTypeInfo)));
            var result = Marshal.GetTypeForITypeInfo(pTypeInfo);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            return result;
        }
        
        internal static string GetTypeLibName(ITypeLib typelib)
        {
            Trace(TraceAction.Entering);
            var result = Marshal.GetTypeLibName(typelib);
            Trace(TraceAction.Leaving, args: (nameof(result), $"`{result}`"));
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
            Assert(() =>
            {
                Debug.Assert(ptr != IntPtr.Zero, "Null pointer passed in");
            });

            Trace(TraceAction.Entering, args: (nameof(ptr), FormatPtr(ptr)));
            var result = Marshal.PtrToStringBSTR(ptr);
            Trace(TraceAction.Leaving, args: (nameof(result), $"`{result}`"));
            return result;
        }

        internal static object PtrToStructure(IntPtr ptr, Type T)
        {
            Assert(() => 
            {
                Debug.Assert(ptr != IntPtr.Zero, "Null pointer passed in");
            });

            Trace(TraceAction.Entering, args: (nameof(ptr), FormatPtr(ptr)));
            var result = Marshal.PtrToStructure(ptr, T);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            return result;
        }

        internal static int QueryInterface(IntPtr pUnk, ref Guid iid, out IntPtr ppv)
        {
            var iid_local = iid;
            Assert(() =>
            {
                Debug.Assert(pUnk != IntPtr.Zero, "Null pointer passed in");
                Debug.Assert(iid_local != Guid.Empty, "Empty IID passed in");
            });

            Trace(TraceAction.Entering, pUnk, args: (nameof(iid), iid.ToString()));
            var result = Marshal.QueryInterface(pUnk, ref iid, out ppv);
            Trace(TraceAction.Leaving, pUnk, args: new (string Name, object Value)[]
            {
                (nameof(result), result),
                (nameof(ppv), ppv)
            });
            PrintRefCount(ppv);
            return result;
        }

        internal static int ReadInt32(IntPtr ptr)
        {
            Assert(() =>
            {
                Debug.Assert(ptr != IntPtr.Zero);
            });

            Trace(TraceAction.Entering, args:(nameof(ptr), FormatPtr(ptr)));
            var result = Marshal.ReadInt32(ptr);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            return result;
        }

        internal static IntPtr ReadIntPtr(IntPtr ptr)
        {
            Assert(() =>
            {
                Debug.Assert(ptr != IntPtr.Zero);
            });

            Trace(TraceAction.Entering, args: (nameof(ptr), FormatPtr(ptr)));
            var result = Marshal.ReadIntPtr(ptr);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            return result;
        }

        internal static int Release(IntPtr pUnk)
        {
            Trace(TraceAction.Entering, pUnk);
            var result = Marshal.Release(pUnk);
            Assert(() =>
            {
                Debug.Assert(result >= 0, "The ref count is negative which is invalid.");
            });
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            PrintRefCount(pUnk, refCount: result);
            return result;
        }

        internal static int ReleaseComObject(object o)
        {
            Assert(() =>
            {
                var ptr = Marshal.GetIUnknownForObject(o);
                var debugResult = Marshal.Release(ptr);
                Debug.Assert(debugResult > 0,
                    $"The ref count is at zero or is invalid before calling the {nameof(Marshal.ReleaseComObject)}.");
            });
            var result = Marshal.ReleaseComObject(o);
            Assert(() =>
            {
                Debug.Assert(result >= 0,
                    $"The ref count is invalid after calling the {nameof(Marshal.ReleaseComObject)}");
            });
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            PrintRefCount(o, result);
            return result;
        }

        internal static int SizeOf(Type t)
        {
            Trace(TraceAction.Entering, args: (nameof(t), t.Name));
            var result = Marshal.SizeOf(t);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            return result;
        }

        internal static int SizeOf(object structure)
        {
            Trace(TraceAction.Entering, args: (nameof(structure), structure.GetType().Name));
            var result = Marshal.SizeOf(structure);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            return result;
        }

        internal static IntPtr StringToBSTR(string s)
        {
            Trace(TraceAction.Entering, args: (nameof(s), $"`{s}`"));
            var result = Marshal.StringToBSTR(s);
            Trace(TraceAction.Leaving, args: (nameof(result), result));
            PrintAlloc(result);
            return result;
        }

        internal static void StructureToPtr<T>(T structure, IntPtr ptr, bool fDeleteOld)
        {
            Trace(TraceAction.Executing, args: new (string Name, object Value)[]
            {
                (nameof(structure), structure.GetType().Name), 
                (nameof(ptr), FormatPtr(ptr)), 
                (nameof(fDeleteOld), fDeleteOld)
            });
            Marshal.StructureToPtr(structure, ptr, fDeleteOld);
            Debug.Assert(fDeleteOld, $"Warning: {nameof(StructureToPtr)} had flag {fDeleteOld} passed with false. That may cause memory leaks.");
        }

        internal static void WriteInt32(IntPtr ptr, int val)
        {
            Assert(() =>
            {
                Debug.Assert(ptr != IntPtr.Zero, "Null pointer passed in");
            });

            Trace(TraceAction.Executing, args: new (string Name, object Value)[]
            {
                (nameof(ptr), FormatPtr(ptr)),
                (nameof(val), val)
            });
            Marshal.WriteInt32(ptr, val);
        }

        internal static void WriteIntPtr(IntPtr ptr, IntPtr val)
        {
            Assert(() =>
            {
                Debug.Assert(ptr != IntPtr.Zero, "Null pointer passed in"); 
            });
            
            Trace(TraceAction.Executing, args: new (string Name, object Value)[]
            {
                (nameof(ptr), FormatPtr(ptr)), 
                (nameof(val), val)
            });
            Marshal.WriteIntPtr(ptr, val);
        }

        //Helper functions
        private enum TraceAction
        {
            Entering,
            Leaving,
            Executing
        }

        [Conditional("TRACE_MARSHAL")]
        private static void Trace(TraceAction action, IntPtr pUnk, [CallerMemberName]string MethodName = null, params (string Name, object Value)[] args)
        {
            Debug.Assert(pUnk != IntPtr.Zero, "Null pointer passed in");
            var argPrint = string.Empty;
            if (args.Length != 0)
            {
                argPrint = args.Aggregate(argPrint, (current, arg) => current + $"; {arg.Name}: {arg.Value}");
            }

            Debug.Print($"{Enum.GetName(typeof(TraceAction), action)} {MethodName}; {nameof(pUnk)}: {FormatPtr(pUnk)} {argPrint}");
        }

        [Conditional("TRACE_MARSHAL")]
        private static void Trace(TraceAction action, [CallerMemberName]string MethodName = null, params (string Name, object Value)[] args)
        {
            var argPrint = string.Empty;
            if (args.Length != 0)
            {
                argPrint = args.Aggregate(argPrint, (current, arg) => current + $"; {arg.Name}: {arg.Value}");
            }

            Debug.Print($"{Enum.GetName(typeof(TraceAction), action)} {MethodName}; {argPrint}");
        }

        [Conditional("TRACE_MARSHAL")]
        private static void Assert(Action assert)
        {
            assert.Invoke();
        }

        [Conditional("REF_COUNT")]
        private static void PrintRefCount(object o, int result)
        {
            Debug.Print($"{nameof(ReleaseComObject)}:: COM Object: {o.GetType().Name} ref count {result}");
        }

        [Conditional("REF_COUNT")]
        private static void PrintRefCount(IntPtr pUnk, [CallerMemberName] string methodName = null, int refCount = 0)
        {
            if (pUnk != IntPtr.Zero)
            {
                if (refCount != 0)
                {
                    Debug.Print($"{nameof(Release)}:: COM Object: {FormatPtr(pUnk)} ref count {refCount}");
                }
                else
                {
                    refCount = Marshal.AddRef(pUnk) - 1;
                    Marshal.Release(pUnk);
                    Debug.Print($"{methodName}:: COM Object: {FormatPtr(pUnk)} ref count {refCount}");
                }
            }
            else
            {
                Debug.Print($"{methodName}:: COM Object: not allocated");
            }
        }

        [Conditional("REF_COUNT")]
        private static void PrintAlloc(IntPtr pUnmanaged, [CallerMemberName] string methodName = null)
        {
            Debug.Print($"{methodName}:: Unmanaged pointer allocated: {FormatPtr(pUnmanaged)}");
        }

        [Conditional("REF_COUNT")]
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
