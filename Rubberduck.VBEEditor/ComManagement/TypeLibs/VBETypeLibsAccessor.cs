using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// An internal representation of the VBE References collection object, as returned from VBE.ActiveVBProject.References, or similar
    /// These offsets are known to be valid across 32-bit and 64-bit versions of VBA and VB6, right back from when VBA6 was first released.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct VBEReferencesObj
    {
        IntPtr _vTable1;     // _References vtable
        IntPtr _vTable2;
        IntPtr _vTable3;
        IntPtr _object1;
        IntPtr _object2;
        public IntPtr _typeLib;
        IntPtr _placeholder1;
        IntPtr _placeholder2;
        IntPtr _refCount;
    }

    /// <summary>
    /// The root class for hooking into the live <see cref="ITypeLib"/>s provided by the VBE
    /// </summary>
    /// <remarks>
    /// WARNING: when using <see cref="VBETypeLibsAccessor"/> directly, do not cache it
    ///   The VBE provides LIVE type library information, so consider it a snapshot at that very moment when you are dealing with it
    ///   Make sure you call VBETypeLibsAccessor.Dispose() as soon as you have done what you need to do with it.
    ///   Once control returns back to the VBE, you must assume that all the ITypeLib/ITypeInfo pointers are now invalid.
    /// </remarks>
    internal class VBETypeLibsAccessor : DisposableList<ITypeLibWrapper>
    {
        internal VBETypeLibsAccessor(IVBE ide)
        {
            // We need at least one project in the VBE.VBProjects collection to be accessible (i.e. unprotected)
            // in order to get access to the list of loaded project TypeLibs using this method
            using (var projects = ide.VBProjects)
            {
                foreach (var project in projects)
                {
                    using (project)
                    {
                        try
                        {
                            using (var references = project.References)
                            {
                                // Now we've got the references object, we can read the internal object structure to grab the ITypeLib
                                var internalReferencesObj =
                                    StructHelper.ReadComObjectStructure<VBEReferencesObj>(references.Target);

                                // Now we've got this one internalReferencesObj.typeLib, we can iterate through ALL loaded project TypeLibs
                                using (var typeLibIterator = new VBETypeLibsIterator(internalReferencesObj._typeLib))
                                {
                                    foreach (var typeLib in typeLibIterator)
                                    {
                                        Add(typeLib);
                                    }
                                }
                            }

                            // we only need access to a single VBProject References object to make it work, so we can return now.
                            return;
                        }
                        catch
                        {
                            // probably a protected project, just move on to the next project.
                        }
                    }
                }
            }

            // return an empty list on error
        }

        internal ITypeLibWrapper Find(string searchLibName)
        {
            foreach (var typeLib in this)
            {
                if (typeLib.Name == searchLibName)
                {
                    return typeLib;
                }
            }
            return null;
        }

        internal ITypeLibWrapper Get(string searchLibName)
        {
            var retVal = Find(searchLibName);
            if (retVal == null)
            {
                throw new ArgumentException($"TypeLibWrapper::Get failed. '{searchLibName}' component not found.");
            }
            return retVal;
        }
    }
}
