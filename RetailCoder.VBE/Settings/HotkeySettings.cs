namespace Rubberduck.Settings
{
    public interface IHotkeySettings
    {
        HotkeySetting[] Settings { get; set; }
    }

    public class HotkeySettings : IHotkeySettings
    {
        public HotkeySetting[] Settings { get; set; }

        public HotkeySettings()
        {
            Settings = new[]
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
                new HotkeySetting{Name=RubberduckHotkey.RefactorExtractMethod.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="M" }
            };
        }
    }
}
