using Rubberduck.AutoComplete;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Xml.Serialization;

namespace Rubberduck.Settings
{
    public interface IAutoCompleteSettings
    {
        HashSet<AutoCompleteSetting> Settings { get; set; }
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class AutoCompleteSettings : IAutoCompleteSettings
    {
        private readonly IEnumerable<AutoCompleteSetting> _defaultSettings;
        private HashSet<AutoCompleteSetting> _settings = new HashSet<AutoCompleteSetting>();

        public AutoCompleteSettings() { /* default constructor required for XML serialization */ }

        public AutoCompleteSettings(IEnumerable<AutoCompleteSetting> defaultSettings)
        {
            _defaultSettings = defaultSettings;
            _settings = defaultSettings.ToHashSet();
        }

        [XmlArrayItem("AutoComplete", IsNullable = false)]
        public HashSet<AutoCompleteSetting> Settings
        {
            get => _settings;
            set
            {
                // Enable loading user settings during deserialization
                if (_defaultSettings == null)
                {
                    if (value != null)
                    {
                        AddUnique(value);
                    }

                    return;
                }

                var defaults = _defaultSettings.ToArray();

                if (value == null || value.Count == 0)
                {
                    _settings = new HashSet<AutoCompleteSetting>(defaults);
                    return;
                }

                _settings = new HashSet<AutoCompleteSetting>();

                var incoming = value;
                AddUnique(incoming);

                //Merge any hotkeys that weren't found in the input.
                foreach (var setting in defaults.Where(setting => _settings.FirstOrDefault(s => s.Key.Equals(setting.Key)) == null))
                {
                    setting.IsEnabled &= !IsDuplicate(setting);
                    _settings.Add(setting);
                }
            }

        }

        private void AddUnique(IEnumerable<AutoCompleteSetting> settings)
        {
            //Only take the first setting if multiple definitions are found.
            foreach (var setting in settings.GroupBy(s => s.Key).Select(autocomplete => autocomplete.First()))
            {
                //Only allow one hotkey to be enabled with the same key combination.
                setting.IsEnabled &= !IsDuplicate(setting);
                _settings.Add(setting);
            }
        }

        private bool IsDuplicate(AutoCompleteSetting setting)
        {
            return _settings.FirstOrDefault(s => s.Key == setting.Key) != null;
        }

        public AutoCompleteSetting GetSetting<TAutoComplete>() where TAutoComplete : IAutoComplete
        {
            return Settings.FirstOrDefault(s => typeof(TAutoComplete).Name.Equals(s.Key))
                ?? GetSetting(typeof(TAutoComplete));
        }

        public AutoCompleteSetting GetSetting(Type autoCompleteType)
        {
            try
            {
                var existing = Settings.FirstOrDefault(s => autoCompleteType.Name.Equals(s.Key));
                if (existing != null)
                {
                    return existing;
                }
                var proto = Convert.ChangeType(Activator.CreateInstance(autoCompleteType, new object[] { null }), autoCompleteType);
                var setting = new AutoCompleteSetting(proto as IAutoComplete);
                Settings.Add(setting);
                return setting;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}