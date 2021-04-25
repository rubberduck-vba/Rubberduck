using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NLog;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// An internal interface supported by VBA for all projects. Obtainable from a VBE hosted ITypeLib 
    /// in order to access a few extra features...
    /// </summary>
    /// <remarks>This internal interface is known to be supported since the very earliest version of VBA6</remarks>
    [ComImport(), Guid("DDD557E0-D96F-11CD-9570-00AA0051E5D4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IVBEProject
    {
        string GetProjectName();                 // same as calling ITypeLib::GetDocumentation(-1)                   
        void SetProjectName(string value);       // same as IVBEProject2::set_ProjectName()
        int GetVbeLCID();
        void Placeholder3();                      // calls IVBEProject2::Placeholder8
        void Placeholder4();
        void Placeholder5();
        void Placeholder6();
        void Placeholder7();
        string GetConditionalCompilationArgs();
        void SetConditionalCompilationArgs(string args);
        void Placeholder8();
        void Placeholder9();
        void Placeholder10();
        void Placeholder11();
        void Placeholder12();
        void Placeholder13();
        int GetReferencesCount();
        IntPtr GetReferenceTypeLib(int referenceIndex);
        void Placeholder16();
        void Placeholder17();
        string GetReferenceString(int referenceIndex); // the raw reference string
        void CompileProject();                            // throws COM exception 0x800A9C64 if error occurred during compile.
    }

    /*
     Not currently used.
    // IVBEProject2, vtable position just before the IVBEProject, not queryable, so needs aggregation
    [ComImport(), Guid("FFFFFFFF-0000-0000-C000-000000000046")]  // 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IVBEProject2
    {
        void Placeholder1();                    // returns E_NOTIMPL
        void SetProjectName(string value);
        void SetProjectVersion(ushort wMajorVerNum, ushort wMinorVerNum);
        void SetProjectGUID(ref Guid value);
        void SetProjectDescription(string value);
        void SetProjectHelpFileName(string value);
        void SetProjectHelpContext(int value);
    }
    */

    /// <summary>
    /// Exposes the VBE specific extensions provided by an <see cref="ITypeLib"/>
    /// </summary>
    internal class TypeLibVBEExtensions : ITypeLibVBEExtensions
    {
        private readonly string _name;
        private readonly IVBEProject _target_IVBEProject;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public ITypeLibReferenceCollection VBEReferences { get; }

        public TypeLibVBEExtensions(ITypeLibWrapper parent, ITypeLibInternal unwrappedTypeLib)
        {
            _name = parent.Name;
            // ReSharper disable once SuspiciousTypeConversion.Global
            // We assume that the caller already has checked the HasVBEExtensions
            // before creating this object for the given type library.
            _target_IVBEProject = (IVBEProject)unwrappedTypeLib;
            VBEReferences = new TypeLibReferenceCollection(this);
        }

        /// <summary>
        /// Silently compiles the whole VBA project represented by this ITypeLib
        /// </summary>
        /// <returns>true if the compilation succeeds</returns>
        public bool CompileProject()
        {
            try
            {
                // FIXME: Prevent an access violation when calling CompileProject(). A easy way to reproduce
                // this AV is to parse an Access project, then do a Compact & Repair, then try to shut down
                // Access. There was an attempt to avoid the AV by calling PlaceHolder3 which did seem to prevent
                // the AV but somehow alters the VBA project in such way that _sometimes_ the user is shown a message
                // that the project has changed and whether the user wants to proceeds. That is a slightly worse fix
                // than the AV prevention, so we had to remove the fix. Reference:
                // https://github.com/rubberduck-vba/Rubberduck/issues/5722
                // https://github.com/rubberduck-vba/Rubberduck/pull/5675

                //try
                //{
                //    _target_IVBEProject.Placeholder3();
                //}
                //catch(Exception ex)
                //{
                //    Logger.Info(ex, $"Cannot compile the VBA project '{_name}' because there may be a potential access violation.");
                //    return false;
                //}

                _target_IVBEProject.CompileProject();
                return true;
            }
            catch (Exception e)
            {
                ThrowOnUnrecongizedCompilerError(e);
                return false;
            }
        }

        [Conditional("DEBUG")]
        private static void ThrowOnUnrecongizedCompilerError(Exception e)
        {
            if (e.HResult != (int) KnownComHResults.E_VBA_COMPILEERROR)
            {
                // this is for debug purposes, to see if the compiler ever returns other errors on failure
                throw new InvalidOperationException("Unrecognized VBE compiler error: \n" + e.ToString());
            }
        }

        /// <summary>
        /// Exposes the raw conditional compilation arguments defined in the VBA project represented by this ITypeLib
        /// format:  "foo = 1 : bar = 2"
        /// </summary>
        public string ConditionalCompilationArgumentsRaw
        {
            get => _target_IVBEProject.GetConditionalCompilationArgs();

            set => _target_IVBEProject.SetConditionalCompilationArgs(value);
        }

        /// <summary>
        /// Exposes the conditional compilation arguments defined in the VBA project represented by this ITypeLib
        /// as a dictionary of key/value pairs
        /// </summary>
        public Dictionary<string, short> ConditionalCompilationArguments
        {
            get
            {
                var args = _target_IVBEProject.GetConditionalCompilationArgs();

                if (args.Length <= 0)
                {
                    return new Dictionary<string, short>();
                }

                var argsArray = args.Split(new[] { ':' });
                return argsArray.Select(item => item.Split('=')).ToDictionary(s => s[0].Trim(), s => short.Parse(s[1]));
            }

            set
            {
                var rawArgsString = string.Join(" : ", value.Select(x => x.Key + " = " + x.Value));
                ConditionalCompilationArgumentsRaw = rawArgsString;
            }
        }

        public int GetVBEReferencesCount()
        {
            return _target_IVBEProject.GetReferencesCount();
        }

        public ITypeLibReference GetVBEReferenceByIndex(int index)
        {
            if (index >= _target_IVBEProject.GetReferencesCount())
            {
                throw new ArgumentException($"Specified index not valid for the references collection (reference {index} in project {_name})");
            }

            return new TypeLibReference(this, index, _target_IVBEProject.GetReferenceString(index));
        }
        
        public ITypeLibWrapper GetVBEReferenceTypeLibByIndex(int index)
        {
            if (index >= _target_IVBEProject.GetReferencesCount())
            {
                throw new ArgumentException($"Specified index not valid for the references collection (reference {index} in project {_name})");
            }

            var referenceTypeLibPtr = _target_IVBEProject.GetReferenceTypeLib(index);
            if (referenceTypeLibPtr == IntPtr.Zero)
            {
                throw new ArgumentException("Reference TypeLib not available - probably a missing reference.");
            }
            return TypeApiFactory.GetTypeLibWrapper(referenceTypeLibPtr, addRef: false);
        }
        
        public ITypeLibReference GetVBEReferenceByGuid(Guid referenceGuid)
        {
            foreach (var reference in VBEReferences)
            {
                if (reference.GUID == referenceGuid)
                {
                    return reference;
                }
            }

            throw new ArgumentException($"Specified GUID not found in references collection {referenceGuid}.");
        }

        /*
         This is not yet used, but here in case we want to use this interface at some point.
        private IVBEProjectEx2 target_IVBEProject2
        {
            get
            {
                if (_cachedIVBProjectEx2 == null)
                {
                    if (HasVBEExtensions)
                    {
                        // This internal VBE interface doesn't have a queryable IID.  
                        // The vtable for this interface directly preceeds the _IVBProjectEx, and we can access it through an aggregation helper
                        var objIVBProjectExPtr = RdMarshal.GetComInterfaceForObject(_wrappedObject, typeof(IVBEProject));
                        _cachedIVBProjectEx2 = ComHelper.ForceComObjPtrToInterfaceViaAggregation<IVBEProject2>(objIVBProjectExPtr - IntPtr.Size, queryForType: false);
                    }
                    else
                    {
                        throw new ArgumentException("This ITypeLib is not hosted by the VBE, so does not support _IVBProjectEx");
                    }
                }

                return (IVBEProject2)_cachedIVBProjectEx2.WrappedObject;
            }
        }*/
    }
}
