using System.Collections.Generic;
using Rubberduck.Inspections;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Settings
{
    public interface IToDoListConfigProvider
    {
        ToDoListSettings Create();
        ToDoListSettings CreateDefaults();
        void Save(ToDoListSettings settings);
    }

    public class ToDoListConfigProvider : IToDoListConfigProvider
    {
        private readonly IPersistanceService<ToDoListSettings> _persister;

        public ToDoListConfigProvider(IPersistanceService<ToDoListSettings> persister)
        {
            _persister = persister;
        }

        public ToDoListSettings Create()
        {
            var prototype = new ToDoListSettings();
            return _persister.Load(prototype) ?? prototype;
        }

        public ToDoListSettings CreateDefaults()
        {
            return new ToDoListSettings();
        }

        public void Save(ToDoListSettings settings)
        {
            _persister.Save(settings);
        }
    }
}
