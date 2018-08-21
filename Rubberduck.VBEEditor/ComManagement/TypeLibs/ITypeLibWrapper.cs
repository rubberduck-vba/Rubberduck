using System;
using System.Collections.Generic;
using Rubberduck.VBEditor.ComManagement.TypeLibsSupport;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    public interface ITypeLibWrapper: System.Runtime.InteropServices.ComTypes.ITypeLib, IDisposable
    {
        string Name { get; }
        string DocString { get; }
        int HelpContext { get; }
        string HelpFile { get; }
        bool HasVBEExtensions { get; }
        int TypesCount { get; }

        TypeInfosCollection TypeInfos { get; }

        System.Runtime.InteropServices.ComTypes.TYPELIBATTR Attributes { get; }

        /// <summary>
        /// Exposes the raw conditional compilation arguments defined in the BA project represented by this ITypeLib
        /// format:  "foo = 1 : bar = 2"
        /// </summary>
        string ConditionalCompilationArgumentsRaw { get; set; }

        /// <summary>
        /// Exposes the conditional compilation arguments defined in the BA project represented by this ITypeLib
        /// as a dictionary of key/value pairs
        /// </summary>
        Dictionary<string, short> ConditionalCompilationArguments { get; set; }

        int GetVBEReferencesCount();
        TypeInfoReference GetVBEReferenceByIndex(int index);
        TypeLibWrapper GetVBEReferenceTypeLibByIndex(int index);
        TypeInfoReference GetVBEReferenceByGuid(Guid referenceGuid);
        TypeInfoWrapper GetSafeTypeInfoByIndex(int index);

        /// <summary>
        /// Silently compiles the whole VBA project represented by this ITypeLib
        /// </summary>
        /// <returns>true if the compilation succeeds</returns>
        bool CompileProject();

        void Document(StringLineBuilder output);
    }
}