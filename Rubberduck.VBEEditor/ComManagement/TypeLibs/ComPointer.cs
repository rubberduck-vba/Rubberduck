using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    internal sealed class ComPointer<TComInterface> : IDisposable
    {
        private readonly IntPtr _pUnk;
        private readonly bool _addRef;
        
        internal static ComPointer<TComInterface> GetObject(IntPtr pUnk, bool addRef)
        {
            return new ComPointer<TComInterface>(pUnk, addRef);
        }

        internal static ComPointer<TComInterface> GetObjectViaAggregation(IntPtr pUnk, bool addRef, bool queryType)
        {
            var comInterface = ComHelper.ComCastViaAggregation<TComInterface>(pUnk, queryType);
            return new ComPointer<TComInterface>(pUnk, comInterface);
        }

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

#if DEBUG
            Debug.Print($"ComPointer:: Created from pointer: pUnk: {RdMarshal.FormatPtr(_pUnk)} interface: {Interface.GetType().Name} - {Interface.GetHashCode()} addRef: {_addRef} refCount: {refCount}");
#endif
        }

        private ComPointer(IntPtr pUnk, TComInterface comInterface)
        {
            _pUnk = pUnk;
            Interface = comInterface;

#if DEBUG
            Debug.Print($"ComPointer:: Created from aggregation: pUnk: {RdMarshal.FormatPtr(_pUnk)} interface: {Interface.GetType().Name} - {Interface.GetHashCode()} addRef: {_addRef}");
#endif
        }

        private ComPointer(TComInterface comInterface)
        {
            Interface = comInterface;
            _pUnk = RdMarshal.GetIUnknownForObject(Interface);

#if DEBUG
            Debug.Print($"ComPointer:: Created from object: pUnk: {RdMarshal.FormatPtr(_pUnk)} interface: {Interface.GetType().Name} - {Interface.GetHashCode()} addRef: {_addRef}");
#endif
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
            var refCount = -1;
            
            if (_addRef)
            {
                refCount = RdMarshal.Release(_pUnk);
            } 

#if DEBUG
            Debug.Print($"ComPointer:: Disposed: _pUnk: {RdMarshal.FormatPtr(_pUnk)} _interface: {Interface.GetType().Name} - {Interface.GetHashCode()} addRef: {_addRef} rcwCount: {rcwCount} refCount: {refCount}");
#endif
            _disposed = true;
        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }

        ~ComPointer()
        {
            Debug.Print("ComPointer:: Finalize called");
            ReleaseUnmanagedResources();
        }
    }
}
