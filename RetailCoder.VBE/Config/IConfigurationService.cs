using System;
using System.Runtime.InteropServices;

namespace Rubberduck.Config
{
    [ComVisible(false)]
    public interface IConfigurationService
    {
        CodeInspection[] GetDefaultCodeInspections();
        Configuration GetDefaultConfiguration();
        ToDoMarker[] GetDefaultTodoMarkers();
        System.Collections.Generic.IList<Rubberduck.Inspections.IInspection> GetImplementedCodeInspections();
        System.Collections.Generic.List<Rubberduck.VBA.Parser.Grammar.ISyntax> GetImplementedSyntax();
        Configuration LoadConfiguration();
        void SaveConfiguration<T>(T toSerialize);
    }
}
