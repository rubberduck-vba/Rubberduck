using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common.Hotkeys;

namespace Rubberduck.Settings
{
    public interface IHotkeySettings
    {
        HotkeySetting[] Settings { get; set; }
    }

    public class HotkeySettings : IHotkeySettings, IEquatable<HotkeySettings>
    {
        private readonly IEnumerable<HotkeySetting> _defaultSettings;
        private HashSet<HotkeySetting> _settings;

        public HotkeySetting[] Settings
        {
            get => _settings?.ToArray();
            set
            {
                // Enable loading user settings during deserialization
                if (_defaultSettings == null)
                {
                    _settings = value == null ? new HashSet<HotkeySetting>() : new HashSet<HotkeySetting>(value);
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

                //Only take the first setting if multiple definitions are found.
                foreach (var setting in incoming.GroupBy(s => s.CommandTypeName).Select(hotkey => hotkey.First()))
                {
                    //Only allow one hotkey to be enabled with the same key combination.
                    setting.IsEnabled &= !IsDuplicate(setting);
                    _settings.Add(setting);
                }

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
        }

        private bool IsDuplicate(HotkeySetting candidate)
        {
            return _settings.FirstOrDefault(
                s =>
                    s.Key1 == candidate.Key1 &&
                    s.Key2 == candidate.Key2 &&
                    s.HasAltModifier == candidate.HasAltModifier &&
                    s.HasCtrlModifier == candidate.HasCtrlModifier &&
                    s.HasShiftModifier == candidate.HasShiftModifier) != null;
        }

        public bool Equals(HotkeySettings other)
        {
            return other != null && Settings.SequenceEqual(other.Settings);
        }

        private static bool IsValid(HotkeySetting candidate)
        {
            //This feels a bit sleazy...
            try
            {
                // ReSharper disable once UnusedVariable
                var test = new Hotkey(new IntPtr(), candidate.ToString(), null);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
