using System;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public enum PROCKIND
    {
        PROCKIND_PROC,
        PROCKIND_LET,
        PROCKIND_SET,
        PROCKIND_GET
    }

    public interface ITypeInfoFunction : IDisposable
    {
        System.Runtime.InteropServices.ComTypes.FUNCDESC FuncDesc { get; }
        string[] NamesArray { get; }
        int NamesArrayCount { get; }
        int MemberID { get; }
        System.Runtime.InteropServices.ComTypes.FUNCFLAGS MemberFlags { get; }
        System.Runtime.InteropServices.ComTypes.INVOKEKIND InvokeKind { get; }
        string Name { get; }
        int ParamCount { get; }
        PROCKIND ProcKind { get; }
        //void Dispose();
    }
}