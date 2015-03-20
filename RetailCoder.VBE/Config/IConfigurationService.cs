using System.Collections.Generic;
using Rubberduck.Inspections;

namespace Rubberduck.Config
{
    public interface IConfigurationService
    {
        CodeInspectionSetting[] GetDefaultCodeInspections();
        Configuration GetDefaultConfiguration();
        ToDoMarker[] GetDefaultTodoMarkers();
        IList<IInspection> GetImplementedCodeInspections();
        Configuration LoadConfiguration();
        void SaveConfiguration<T>(T toSerialize);
    }
}
