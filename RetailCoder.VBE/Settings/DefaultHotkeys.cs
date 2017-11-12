using System.Collections.Generic;
using System.Linq;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.Settings
{
    public class DefaultHotkeys
    {
        public IEnumerable<HotkeySetting> Hotkeys { get; }

        public DefaultHotkeys(IEnumerable<CommandBase> commands)
        {
            var hotkeys = new List<HotkeySetting>();

            var command = commands.FirstOrDefault(c => c.GetType() == typeof(CodePaneRefactorRenameCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    HasShiftModifier = true,
                    Key1 = "R"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(RefactorEncapsulateFieldCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    HasShiftModifier = true,
                    Key1 = "F"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(RefactorExtractMethodCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    HasShiftModifier = true,
                    Key1 = "M"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(RefactorMoveCloserToUsageCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    HasShiftModifier = true,
                    Key1 = "C"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(CodeExplorerCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    Key1 = "C"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(ExportAllCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    HasShiftModifier = true,
                    Key1 = "E"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(FindSymbolCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    Key1 = "T"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(IndentCurrentModuleCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    Key1 = "M"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(IndentCurrentProcedureCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    Key1 = "P"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(InspectionResultsCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    HasShiftModifier = true,
                    Key1 = "I"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(ReparseCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    Key1 = "`"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(SourceControlCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    HasShiftModifier = true,
                    Key1 = "D6"
                });
            }
            command = commands.FirstOrDefault(c => c.GetType() == typeof(TestExplorerCommand));
            if (command != null)
            {
                hotkeys.Add(new HotkeySetting(command)
                {
                    IsEnabled = true,
                    HasCtrlModifier = true,
                    HasShiftModifier = true,
                    Key1 = "T"
                });
            }

            Hotkeys = hotkeys;
        }
    }
}
