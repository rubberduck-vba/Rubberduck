using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common.Hotkeys;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Settings
{
    public interface IHotkeySettings
    {
        HotkeySetting[] Settings { get; set; }
    }

    public class HotkeySettings : IHotkeySettings, IEquatable<HotkeySettings>
    {
        private readonly IEnumerable<HotkeySetting> _defaultSettings;
        private HashSet<HotkeySetting> _settings = new HashSet<HotkeySetting>();

        public HotkeySetting[] Settings
        {
            get => _settings?.ToArray();
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

                if (value == null || value.Length == 0)
                {
                    _settings = new HashSet<HotkeySetting>(defaults);
                    return;
                }

                _settings = new HashSet<HotkeySetting>();

                var incoming = value.ToList();
                //Make sure settings are valid to keep trash out of the config file.
                var hotkeyCommandTypeNames = defaults.Select(h => h.CommandTypeName);
                incoming.RemoveAll(h => !hotkeyCommandTypeNames.Contains(h.CommandTypeName) || !IsValid(h));

                AddUnique(incoming);

                //Merge any hotkeys that weren't found in the input.
                foreach (var setting in defaults.Where(setting => _settings.FirstOrDefault(s => s.CommandTypeName.Equals(setting.CommandTypeName)) == null))
                {
                    setting.IsEnabled &= !IsDuplicate(setting);
                    _settings.Add(setting);
                }
            }
        }

        /// <Summary>
        /// Default constructor required for XML serialization.
        /// </Summary>
        public HotkeySettings()
        {
        }

        public HotkeySettings(IEnumerable<HotkeySetting> defaultSettings)
        {
            _defaultSettings = defaultSettings;
            _settings = defaultSettings.ToHashSet();
        }

        public bool Equals(HotkeySettings other)
        {
            return other != null && Settings.SequenceEqual(other.Settings);
        }

        private static bool IsValid(HotkeySetting candidate)
        {
            return candidate.IsValid;
        }
    }
}
