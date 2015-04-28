using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace Rubberduck.Config
{
    public interface IConfigurationService<T>
    {
        T LoadConfiguration();
        void SaveConfiguration(T toSerialize);
    }

    public interface IGeneralConfigService : IConfigurationService<Configuration>
    {
        CodeInspectionSetting[] GetDefaultCodeInspections();
        Configuration GetDefaultConfiguration();
        ToDoMarker[] GetDefaultTodoMarkers();
        IList<Rubberduck.Inspections.IInspection> GetImplementedCodeInspections();
    }

    //todo: define source control config and inherit from IConfigurationService<SourceControlConfig>
    public interface ISourceControlConfigService
    {

    }
}
