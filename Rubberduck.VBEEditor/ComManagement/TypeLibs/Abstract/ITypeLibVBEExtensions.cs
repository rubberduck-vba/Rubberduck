using System;
using System.Collections.Generic;

namespace Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeLibVBEExtensions
    {
        /// <summary>
        /// Silently compiles the whole VBA project represented by this ITypeLib
        /// </summary>
        /// <returns>true if the compilation succeeds</returns>
        bool CompileProject();

        /// <summary>
        /// Exposes the raw conditional compilation arguments defined in the VBA project represented by this ITypeLib
        /// format:  "foo = 1 : bar = 2"
        /// </summary>
        string ConditionalCompilationArgumentsRaw { get; set; }

        /// <summary>
        /// Exposes the conditional compilation arguments defined in the VBA project represented by this ITypeLib
        /// as a dictionary of key/value pairs
        /// </summary>
        Dictionary<string, short> ConditionalCompilationArguments { get; set; }

        int GetVBEReferencesCount();
        ITypeLibReference GetVBEReferenceByIndex(int index);
        ITypeLibWrapper GetVBEReferenceTypeLibByIndex(int index);
        ITypeLibReference GetVBEReferenceByGuid(Guid referenceGuid);
        ITypeLibReferenceCollection VBEReferences { get; }
    }
}