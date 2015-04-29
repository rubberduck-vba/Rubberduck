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
}
