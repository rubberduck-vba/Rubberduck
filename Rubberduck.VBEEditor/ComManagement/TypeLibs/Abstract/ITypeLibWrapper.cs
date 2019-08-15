using System;
using System.Runtime.InteropServices.ComTypes;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeLibWrapper : ITypeLib, IDisposable
    {
        string Name { get; }
        string DocString { get; }
        int HelpContext { get; }
        string HelpFile { get; }
        bool HasVBEExtensions { get; }
        int TypesCount { get; }

        ITypeInfoWrapperCollection TypeInfos { get; }
        ITypeLibVBEExtensions VBEExtensions { get; }

        System.Runtime.InteropServices.ComTypes.TYPELIBATTR Attributes { get; }

        int GetSafeTypeInfoByIndex(int index, out ITypeInfoWrapper outTI);
    }
}