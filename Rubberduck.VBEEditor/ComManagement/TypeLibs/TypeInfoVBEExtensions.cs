using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// An internal interface exposed by VBA for all components (modules, class modules, etc)
    /// </summary>
    /// <remarks>This internal interface is known to be supported since the very earliest version of VBA6</remarks>
    [ComImport(), Guid("DDD557E1-D96F-11CD-9570-00AA0051E5D4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IVBEComponent
    {
        void Placeholder1();
        void Placeholder2();
        void Placeholder3();
        void Placeholder4();
        void Placeholder5();
        void Placeholder6();
        void Placeholder7();
        void Placeholder8();
        void Placeholder9();
        void Placeholder10();
        void Placeholder11();
        void Placeholder12();
        void CompileComponent();
        void Placeholder14();
        IDispatch GetStdModAccessor();
        void Placeholder16();
        void Placeholder17();
        void Placeholder18();
        void Placeholder19();
        void Placeholder20();
        void Placeholder21();
        void Placeholder22();
        void Placeholder23();
        void Placeholder24();
        void Placeholder25();
        void Placeholder26();
        void Placeholder27();
        void Placeholder28();
        void Placeholder29();
        void Placeholder30();
        void Placeholder31();
        void Placeholder32();
        void Placeholder33();
        void GetSomeRelatedTypeInfoPtrs(out IntPtr a, out IntPtr b);        // returns 2 TypeInfos, seemingly related to this ITypeInfo, but slightly different.
    }

    /// <summary>
    /// Exposes the VBE specific extensions provided by an ITypeInfo
    /// </summary>
    internal class TypeInfoVBEExtensions : ITypeInfoVBEExtensions, IDisposable
    {
        private readonly ITypeInfoWrapper _parent;
        //private readonly IVBEComponent _target_IVBEComponent;
        private readonly ComPointer<IVBEComponent> _vbeComponentPointer;
        private IVBEComponent _vbeComponent => _vbeComponentPointer.Interface;

        public TypeInfoVBEExtensions(ITypeInfoWrapper parent, IntPtr tiPtr)
        {
            _parent = parent;
            //_target_IVBEComponent = ComHelper.ComCastViaAggregation<IVBEComponent>(tiPtr);
            _vbeComponentPointer = ComPointer<IVBEComponent>.GetObjectViaAggregation(tiPtr, false, true);
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _isDisposed = true;

            // We shouldn't dispose the containing ITypeInfoWrapper, since it is required
            // to create the class with it as a parameter and ITypeInfoWrapper should be
            // the one to dispose of this class. However, we did create the IVBEComponent,
            // so we'll dispose of it here.
            _vbeComponentPointer?.Dispose();

            /*if (_target_IVBEComponent != null)
            {
                RdMarshal.ReleaseComObject(_target_IVBEComponent);
            }*/
        }

        /// <summary>
        /// Silently compiles the individual VBE component (class/module etc)
        /// </summary>
        /// <returns>true if this module, plus any direct dependent modules compile successfully</returns>
        public bool CompileComponent()
        {
            try
            {
                //_target_IVBEComponent.CompileComponent();
                _vbeComponent.CompileComponent();
                return true;
            }
            catch (Exception e)
            {
                ThrowOnUnrecongizedCompileError(e);

                return false;
            }
        }

        [Conditional("DEBUG")]
        private static void ThrowOnUnrecongizedCompileError(Exception e)
        {
            if (e.HResult != (int) KnownComHResults.E_VBA_COMPILEERROR)
            {
                // When debugging we want to know if there are any other errors returned by the compiler as
                // the error code might be useful.
                throw new ArgumentException("Unrecognized VBE compiler error: \n" + e.ToString());
            }
        }

        /// <summary>
        /// Provides an accessor object for invoking methods on a standard module in a VBA project
        /// </summary>
        /// <remarks>caller is responsible for calling ReleaseComObject</remarks>
        /// <returns>the accessor object</returns>
        public IDispatch GetStdModAccessor()
        {
            //return _target_IVBEComponent.GetStdModAccessor();
            return _vbeComponent.GetStdModAccessor();
        }

        /// <summary>
        /// Executes a procedure inside a standard module in a VBA project
        /// </summary>
        /// <param name="name">the name of the procedure to invoke</param>
        /// <param name="args">arguments to pass to the procedure</param>
        /// <remarks>the returned object can be a COM object, and the callee is responsible for releasing it appropriately</remarks>
        /// <returns>an object representing the return value from the procedure, or null if none.</returns>
        public object StdModExecute(string name, object[] args = null)
        {
            // We search for the dispId using the real type info rather than using staticModule.GetIdsOfNames, 
            // as we can then also include PRIVATE scoped procedures.
            var func = _parent.Funcs.Find(name, PROCKIND.PROCKIND_PROC);
            if (func == null)
            {
                throw new ArgumentException($"StdModExecute failed.  Couldn't find procedure named '{name}'");
            }

            var staticModule = GetStdModAccessor();

            try
            {
                return IDispatchHelper.Invoke(staticModule, func.FuncDesc.memid, IDispatchHelper.InvokeKind.DISPATCH_METHOD, args);
            }
            finally
            {
                RdMarshal.ReleaseComObject(staticModule);
            }
        }
    }
}
