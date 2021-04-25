using System;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeInfoVariable : IDisposable
    {
        string Name { get; }
        int MemberID { get; }
        VARFLAGS MemberFlags { get; }
        //void Dispose();
    }
}