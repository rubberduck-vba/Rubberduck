using System;
using System.Diagnostics;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged
{
    /// <summary>
    /// The <see cref="ComPointer{TComInterface}"/> encapsulates the conversion between
    /// the unmanaged IUnknown pointer and the corresponding managed COM interface. Because
    /// unmanaged pointers are technically unsafe, and even though we can use
    /// <see cref="System.Runtime.InteropServices.Marshal"/> without unsafe/fixed blocks, we must
    /// treat all code as unsafe. Furthermore, several methods from the class implicitly does an
    /// IUnknown::AddRef. Further complicating the issue, we sometime may obtain the pointer in
    /// a manner where we must add the reference ourselves to ensure the referenced object does not
    /// fall out of scope at wrong time. Thus, the class links the lifetime of both the unmanaged
    /// IUnknown pointer and the managed COM interface (among with its RCW).
    /// </summary>
    /// <typeparam name="TComInterface">
    /// A managed version of COM interface for which the IUnknown pointer must implement.
    /// </typeparam>
    /// <remarks>
    /// The class is meant to be used as a field within another class that provides a COM
    /// interface but requires low-level unmanaged access to the same COM object internally.
    /// It should not be created or passed among classes. Furthermore, the class that uses it
    /// must be <see cref="IDisposable"/> so that it can invoke the <see cref="Dispose"/> on
    /// the field of this type.
    /// </remarks>
    internal sealed class ComPointer<TComInterface> : IDisposable
    {
        private readonly IntPtr _pUnk;
        private readonly bool _addRef;

        /// <summary>
        /// Converts an unmanaged pointer into a <see cref="ComPointer{TComInterface}"/>
        /// </summary>
        /// <param name="pUnk">
        /// Unmanaged IUnknown pointer
        /// </param>
        /// <param name="addRef">
        /// If true, an IUnknown::AddRef will be applied on the pointer. This is required when
        /// the pointer is provided via other methods besides the methods provided in the
        /// <see cref="System.Runtime.InteropServices.Marshal"/>. 
        /// </param>
        /// <returns>
        /// A <see cref="ComPointer{TComInterface}"/> encapsulating the IUnknown pointer and the
        /// managed interface. Use <see cref="Interface"/> to obtain the managed interface.
        /// </returns>
        internal static ComPointer<TComInterface> GetObject(IntPtr pUnk, bool addRef)
        {
            return new ComPointer<TComInterface>(pUnk, addRef);
        }

        /// <summary>
        /// Converts an unmanaged pointer into a <see cref="ComPointer{TComInterface}"/>, using
        /// aggregation rather than type-casting which may fail due to strict rules enforced by
        /// managed code. 
        /// </summary>
        /// <param name="pUnk">
        /// Unmanaged IUnknown pointer
        /// </param>
        /// <param name="addRef">
        /// If true, an IUnknown::AddRef will be applied on the pointer. This is required when
        /// the pointer is provided via other methods besides the methods provided in the
        /// <see cref="System.Runtime.InteropServices.Marshal"/>. 
        /// </param>
        /// <param name="queryType">
        /// Indicate if IUnknown::QueryInterface should be invoked to obtain the interface.
        /// Refer to <see cref="ComHelper.ComCastViaAggregation{T}"/> for details.
        /// </param>
        /// <returns>
        /// A <see cref="ComPointer{TComInterface}"/> encapsulating the IUnknown pointer and the
        /// managed interface. Use <see cref="Interface"/> to obtain the managed interface.
        /// </returns>
        internal static ComPointer<TComInterface> GetObjectViaAggregation(IntPtr pUnk, bool addRef, bool queryType)
        {
            var comInterface = ComHelper.ComCastViaAggregation<TComInterface>(pUnk, queryType);
            return comInterface == null 
                ? null 
                : new ComPointer<TComInterface>(pUnk, comInterface);
        }

        /// <summary>
        /// Converts a managed COM object (e.g. a RCW) into a <see cref="ComPointer{TComInterface}"/> of
        /// the same type. A pointer will be extracted and stored. 
        /// </summary>
        /// <param name="comInterface">
        /// The type of the COM object
        /// </param>
        /// <returns>
        /// A <see cref="ComPointer{TComInterface}"/> encapsulating the IUnknown pointer and the
        /// managed interface. Use <see cref="ExtractPointer"/> to obtain the unmanaged pointer.
        /// </returns>
        internal static ComPointer<TComInterface> GetPointer(TComInterface comInterface)
        {
            return new ComPointer<TComInterface>(comInterface);
        }

        private ComPointer(IntPtr pUnk, bool addRef)
        {
            var refCount = -1;

            _pUnk = pUnk;
            if (addRef)
            {
                _addRef = true;
                refCount = RdMarshal.AddRef(_pUnk); 
            }

            Interface = (TComInterface)RdMarshal.GetTypedObjectForIUnknown(pUnk, typeof(TComInterface));

            ConstructorPointerPrint(refCount);
        }

        [Conditional("TRACE_COM_POINTERS")]
        private void ConstructorPointerPrint(int refCount)
        {
            Debug.Print($"ComPointer:: Created from pointer: pUnk: {RdMarshal.FormatPtr(_pUnk)} interface: {typeof(TComInterface).Name} - {Interface.GetHashCode()} addRef: {_addRef} refCount: {refCount}");
        }

        private ComPointer(IntPtr pUnk, TComInterface comInterface)
        {
            _pUnk = pUnk;
            Interface = comInterface;

            ConstructorAggregatedPrint();
        }

        [Conditional("TRACE_COM_POINTERS")]
        private void ConstructorAggregatedPrint()
        {
            Debug.Print($"ComPointer:: Created from aggregation: pUnk: {RdMarshal.FormatPtr(_pUnk)} interface: {typeof(TComInterface).Name} - {Interface.GetHashCode()} addRef: {_addRef}");
        }

        private ComPointer(TComInterface comInterface)
        {
            Interface = comInterface;
            _pUnk = RdMarshal.GetIUnknownForObject(Interface);

            ConstructorObjectPrint();
        }

        [Conditional("TRACE_COM_POINTERS")]
        private void ConstructorObjectPrint()
        {
            Debug.Print($"ComPointer:: Created from object: pUnk: {RdMarshal.FormatPtr(_pUnk)} interface: {typeof(TComInterface).Name} - {Interface.GetHashCode()} addRef: {_addRef}");
        }

        internal IntPtr ExtractPointer()
        {
            return _pUnk;
        }

        internal TComInterface Interface { get; }

        private bool _disposed;
        private void ReleaseUnmanagedResources()
        {
            if(_disposed) return;

            var rcwCount = RdMarshal.ReleaseComObject(Interface);
            var addRef = _addRef;

            TraceRelease(rcwCount, ref addRef);
            if (addRef)
            {
                RdMarshal.Release(_pUnk);
            }

            _disposed = true;
        }

        [Conditional("TRACE_COM_POINTERS")]
        private void TraceRelease(int rcwCount, ref bool addRef)
        {
            if (!addRef)
            {
                // Temporarily add a ref so that we can safely call IUnknown::Release
                // to report the ref count in the log.
                RdMarshal.AddRef(_pUnk);
            }
            var refCount = RdMarshal.Release(_pUnk);

            Debug.Print($"ComPointer:: Disposed: _pUnk: {RdMarshal.FormatPtr(_pUnk)} _interface: {typeof(TComInterface).Name} - {Interface.GetHashCode()} addRef: {_addRef} rcwCount: {rcwCount} refCount: {refCount}");

            addRef = false;
        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }

        ~ComPointer()
        {
            Debug.Print("ComPointer:: Finalize called -- This is most likely a bug!");
            ReleaseUnmanagedResources();
        }
    }
}
