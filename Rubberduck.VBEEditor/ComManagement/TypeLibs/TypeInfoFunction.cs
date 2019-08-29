using System;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// A class that represents a function definition described in <see cref="ComTypes.FUNCDESC"/>
    /// within a <see cref="ComTypes.ITypeInfo"/>.  
    /// </summary>
    internal class TypeInfoFunction : ITypeInfoFunction
    {
        private readonly ComTypes.ITypeInfo _typeInfo;
        private readonly IntPtr _funcDescPtr;
        private readonly string[] _names = new string[255];   // includes argument names
        private readonly int _cNames = 0;

        public ComTypes.FUNCDESC FuncDesc { get; }

        public string[] NamesArray { get => _names; }
        public int NamesArrayCount { get => _cNames; }
        public int MemberID => FuncDesc.memid;
        public ComTypes.FUNCFLAGS MemberFlags => (ComTypes.FUNCFLAGS)FuncDesc.wFuncFlags;
        public ComTypes.INVOKEKIND InvokeKind => FuncDesc.invkind;


        public TypeInfoFunction(ComTypes.ITypeInfo typeInfo, int funcIndex)
        {
            _typeInfo = typeInfo;

            _typeInfo.GetFuncDesc(funcIndex, out _funcDescPtr);
            FuncDesc = StructHelper.ReadStructureUnsafe<ComTypes.FUNCDESC>(_funcDescPtr);

            typeInfo.GetNames(FuncDesc.memid, _names, _names.Length, out _cNames);
            if (_cNames == 0) _names[0] = "[unnamed]";
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

            if (_funcDescPtr != IntPtr.Zero)
            {
                _typeInfo.ReleaseFuncDesc(_funcDescPtr);
            }

            _isDisposed = true;
        }

        public string Name => _names[0];
        public int ParamCount => FuncDesc.cParams;

        public PROCKIND ProcKind
        {
            get
            {
                // _funcDesc.invkind is a set of flags, and as such we convert into PROCKIND for simplicity
                if (FuncDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUTREF))
                {
                    return PROCKIND.PROCKIND_SET;
                }
                if (FuncDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYPUT))
                {
                    return PROCKIND.PROCKIND_LET;
                }
                if (FuncDesc.invkind.HasFlag(ComTypes.INVOKEKIND.INVOKE_PROPERTYGET))
                {
                    return PROCKIND.PROCKIND_GET;
                }
                return PROCKIND.PROCKIND_PROC;
            }
        }
    }
}
