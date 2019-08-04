using System;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// A class that represents a variable or field described in a <see cref="ComTypes.VARDESC"/>
    /// within a <see cref="ComTypes.ITypeInfo"/>
    /// </summary>
    internal class TypeInfoVariable : ITypeInfoVariable
    {
        private readonly ComTypes.ITypeInfo _typeInfo;
        private readonly ComTypes.VARDESC _varDesc;
        private readonly IntPtr _varDescPtr;

        public string Name { get; private set; }
        public int MemberID => _varDesc.memid;
        public ComTypes.VARFLAGS MemberFlags => (ComTypes.VARFLAGS)_varDesc.wVarFlags;

        public TypeInfoVariable(ComTypes.ITypeInfo typeInfo, int index)
        {
            _typeInfo = typeInfo;

            _typeInfo.GetVarDesc(index, out _varDescPtr);
            _varDesc = StructHelper.ReadStructureUnsafe<ComTypes.VARDESC>(_varDescPtr);

            var names = new string[1];
            typeInfo.GetNames(_varDesc.memid, names, 1, out var actualCount);
            Name = actualCount >= 1 ? names[0] : "[unnamed]";
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            if (_varDescPtr != IntPtr.Zero)
            {
                ((ITypeInfoInternal)_typeInfo).ReleaseVarDesc(_varDescPtr);
            }
            _isDisposed = true;
        }
    }

}
