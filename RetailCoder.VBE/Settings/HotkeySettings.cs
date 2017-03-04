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
        private static readonly HotkeySetting[] Defaults = 
        {
            new HotkeySetting{Name=RubberduckHotkey.ParseAll.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="`" },
            new HotkeySetting{Name=RubberduckHotkey.IndentProcedure.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="P" },
            new HotkeySetting{Name=RubberduckHotkey.IndentModule.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="M" },
            new HotkeySetting{Name=RubberduckHotkey.CodeExplorer.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="R" },
            new HotkeySetting{Name=RubberduckHotkey.FindSymbol.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="T" },
            new HotkeySetting{Name=RubberduckHotkey.InspectionResults.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="I" },
            new HotkeySetting{Name=RubberduckHotkey.TestExplorer.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="T" },
            new HotkeySetting{Name=RubberduckHotkey.RefactorMoveCloserToUsage.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="C" },
            new HotkeySetting{Name=RubberduckHotkey.RefactorRename.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="R" },
            new HotkeySetting{Name=RubberduckHotkey.RefactorExtractMethod.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="M" },
            new HotkeySetting{Name=RubberduckHotkey.SourceControl.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="D6" },
            new HotkeySetting{Name=RubberduckHotkey.RefactorEncapsulateField.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="E" }
        };

        private HashSet<HotkeySetting> _settings;

        public HotkeySettings()
        {
            Settings = Defaults.ToArray();
        }

        public HotkeySetting[] Settings
        {
            get { return _settings.ToArray(); }
            set
            {
                if (value == null || value.Length == 0)
                {
                    _settings = new HashSet<HotkeySetting>(Defaults);
                    return;
                }
                _settings = new HashSet<HotkeySetting>();
                var incoming = value.ToList();
                //Make sure settings are valid to keep trash out of the config file.
                RubberduckHotkey assigned;
                incoming.RemoveAll(h => !Enum.TryParse(h.Name, out assigned) || !IsValid(h));

                //Only take the first setting if multiple definitions are found.
                foreach (var setting in incoming.GroupBy(s => s.Name).Select(hotkey => hotkey.First()))
                {
                    //Only allow one hotkey to be enabled with the same key combination.
                    setting.IsEnabled &= !IsDuplicate(setting);
                    _settings.Add(setting);
                }

                //Merge any hotkeys that weren't found in the input.
                foreach (var setting in Defaults.Where(setting => _settings.FirstOrDefault(s => s.Name.Equals(setting.Name)) == null))
                {
                    setting.IsEnabled &= !IsDuplicate(setting);
                    _settings.Add(setting);
                }
            }
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
