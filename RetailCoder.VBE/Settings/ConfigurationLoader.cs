using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Win32;
using Rubberduck.Inspections;
using Rubberduck.UI;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.Refactorings;
using MessageBox = System.Windows.Forms.MessageBox;

namespace Rubberduck.Settings
{
    public interface IGeneralConfigService : IConfigurationService<Configuration>
    {
        Configuration GetDefaultConfiguration();
    }

    public class ConfigurationLoader : XmlConfigurationServiceBase<Configuration>, IGeneralConfigService
    {
        private readonly IEnumerable<IInspection> _inspections;
        private readonly IEnumerable<ICommand> _commands;

        public ConfigurationLoader(IEnumerable<IInspection> inspections, IEnumerable<ICommand> commands)
        {
            _inspections = inspections;
            _commands = commands;
        }

        protected override string ConfigFile
        {
            get { return Path.Combine(rootPath, "rubberduck.config"); }
        }

        /// <summary>
        /// Loads the configuration from Rubberduck.config xml file.
        /// </summary>
        /// <remarks>
        /// Returns default configuration when an IOException is caught.
        /// </remarks>
        public override Configuration LoadConfiguration()
        {
            //deserialization can silently fail for just parts of the config, 
            //so we null-check and return defaults if necessary.

            var config = base.LoadConfiguration();

            if (config.UserSettings.GeneralSettings == null)
            {
                config.UserSettings.GeneralSettings = GetDefaultGeneralSettings();
            }
            else
            {
                AssociateHotkeyCommands(config);
            }

            // 0 is the default, and parses just fine into a `char`.  We require '.' or '/'.
            if (!new[] {',', '/'}.Contains(config.UserSettings.GeneralSettings.Delimiter))
            {
                config.UserSettings.GeneralSettings.Delimiter = '.';
            }

            if (config.UserSettings.ToDoListSettings == null)
            {
                config.UserSettings.ToDoListSettings = new ToDoListSettings(GetDefaultTodoMarkers());
            }

            if (config.UserSettings.CodeInspectionSettings == null)
            {
                config.UserSettings.CodeInspectionSettings = new CodeInspectionSettings(GetDefaultCodeInspections());
            }

            if (config.UserSettings.UnitTestSettings == null)
            {
                config.UserSettings.UnitTestSettings = new UnitTestSettings();
            }

            if (config.UserSettings.IndenterSettings == null)
            {
                config.UserSettings.IndenterSettings = GetDefaultIndenterSettings();
            }

            var configInspections = config.UserSettings.CodeInspectionSettings.CodeInspections.ToList();
            
            configInspections = MergeImplementedInspectionsNotInConfig(configInspections, _inspections);
            config.UserSettings.CodeInspectionSettings.CodeInspections = configInspections.ToArray();

            return config;
        }

        private IDictionary<RubberduckHotkey, ICommand> GetCommandMappings()
        {
            return new Dictionary<RubberduckHotkey, ICommand>
            {
                { RubberduckHotkey.ParseAll, Command<ReparseCommand>() },
                { RubberduckHotkey.CodeExplorer, Command<CodeExplorerCommand>() },
                { RubberduckHotkey.IndentModule, Command<IndentCurrentModuleCommand>() },
                { RubberduckHotkey.IndentProcedure, Command<IndentCurrentProcedureCommand>() },
                { RubberduckHotkey.FindSymbol, Command<FindSymbolCommand>() },
                { RubberduckHotkey.RefactorMoveCloserToUsage, Command<RefactorMoveCloserToUsageCommand>() },
                { RubberduckHotkey.InspectionResults, Command<InspectionResultsCommand>() },
                { RubberduckHotkey.RefactorExtractMethod, Command<RefactorExtractMethodCommand>() },
                { RubberduckHotkey.RefactorRename, Command<CodePaneRefactorRenameCommand>() },
                { RubberduckHotkey.TestExplorer, Command<TestExplorerCommand>() }
            };
        }

        private ICommand Command<TCommand>() where TCommand : ICommand
        {
            return _commands.OfType<TCommand>().SingleOrDefault();
        }

        private void AssociateHotkeyCommands(Configuration config)
        {
            var mappings = GetCommandMappings();
            foreach (var setting in config.UserSettings.GeneralSettings.HotkeySettings)
            {
                RubberduckHotkey hotkey;
                if (Enum.TryParse(setting.Name, out hotkey))
                {
                    setting.Command = mappings[hotkey];
                    ((CommandBase)setting.Command).ShortcutText = setting.ToMenuHotkeyString(); // yuck
                }
            }
        }

        protected override Configuration HandleIOException(IOException ex)
        {
            return GetDefaultConfiguration();
        }

        protected override Configuration HandleInvalidOperationException(InvalidOperationException ex)
        {
            var folder = Path.GetDirectoryName(ConfigFile);
            var newFilePath = folder + "\\rubberduck.config." + DateTime.UtcNow.ToString().Replace('/', '.').Replace(':', '.') + ".bak";

            var message = string.Format(RubberduckUI.PromptLoadDefaultConfig, ex.Message, ex.InnerException.Message, ConfigFile, newFilePath);
            MessageBox.Show(message, RubberduckUI.LoadConfigError, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);

            using (var fs = File.Create(@newFilePath))
            {
                using (var reader = new StreamReader(folder + "\\rubberduck.config"))
                using (var writer = new StreamWriter(fs, Encoding.UTF8))
                {
                    writer.Write(reader.ReadToEnd());
                }
            }

            var config = GetDefaultConfiguration();
            SaveConfiguration(config);
            return config;
        }

        private List<CodeInspectionSetting> MergeImplementedInspectionsNotInConfig(List<CodeInspectionSetting> configInspections, IEnumerable<IInspection> implementedInspections)
        {
            foreach (var implementedInspection in implementedInspections)
            {
                var inspection = configInspections.SingleOrDefault(i => i.Name == implementedInspection.Name);
                if (inspection == null)
                {
                    configInspections.Add(new CodeInspectionSetting(implementedInspection));
                }
                else
                {
                    // description isn't serialized
                    inspection.Description = implementedInspection.Description;
                }
            }
            return configInspections;
        }

        public Configuration GetDefaultConfiguration()
        {
            var userSettings = new UserSettings(
                                    GetDefaultGeneralSettings(), 
                                    new ToDoListSettings(GetDefaultTodoMarkers()),
                                    new CodeInspectionSettings(GetDefaultCodeInspections()),
                                    new UnitTestSettings(),
                                    GetDefaultIndenterSettings());

            return new Configuration(userSettings);
        }

        private GeneralSettings GetDefaultGeneralSettings()
        {
            var commandMappings = GetCommandMappings();
            return new GeneralSettings(new DisplayLanguageSetting("en-US"),
                new[]
                {
                    new HotkeySetting{Name=RubberduckHotkey.ParseAll.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="`", Command = commandMappings[RubberduckHotkey.ParseAll]},
                    new HotkeySetting{Name=RubberduckHotkey.IndentProcedure.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="P", Command = commandMappings[RubberduckHotkey.IndentProcedure]},
                    new HotkeySetting{Name=RubberduckHotkey.IndentModule.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="M", Command = commandMappings[RubberduckHotkey.IndentModule]},
                    new HotkeySetting{Name=RubberduckHotkey.CodeExplorer.ToString(), IsEnabled=false, HasCtrlModifier = true, Key1="R", Command = commandMappings[RubberduckHotkey.CodeExplorer]},
                    new HotkeySetting{Name=RubberduckHotkey.FindSymbol.ToString(), IsEnabled=true, HasCtrlModifier = true, Key1="T", Command = commandMappings[RubberduckHotkey.InspectionResults]},
                    new HotkeySetting{Name=RubberduckHotkey.InspectionResults.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="I", Command = commandMappings[RubberduckHotkey.InspectionResults]},
                    new HotkeySetting{Name=RubberduckHotkey.TestExplorer.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="T", Command = commandMappings[RubberduckHotkey.TestExplorer]},
                    new HotkeySetting{Name=RubberduckHotkey.RefactorMoveCloserToUsage.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="C", Command = commandMappings[RubberduckHotkey.RefactorMoveCloserToUsage]},
                    new HotkeySetting{Name=RubberduckHotkey.RefactorRename.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="R", Command = commandMappings[RubberduckHotkey.RefactorRename]},
                    new HotkeySetting{Name=RubberduckHotkey.RefactorExtractMethod.ToString(), IsEnabled=true, HasCtrlModifier = true, HasShiftModifier = true, Key1="M", Command = commandMappings[RubberduckHotkey.RefactorExtractMethod]}
                },
                false, 10, '.');
        }

        public ToDoMarker[] GetDefaultTodoMarkers()
        {
            var note = new ToDoMarker(RubberduckUI.TodoMarkerNote);
            var todo = new ToDoMarker(RubberduckUI.TodoMarkerTodo);
            var bug = new ToDoMarker(RubberduckUI.TodoMarkerBug);

            return new[] { note, todo, bug };
        }

        /// <summary>   Converts implemented code inspections into array of Config.CodeInspection objects. </summary>
        /// <returns>   An array of Config.CodeInspection. </returns>
        public CodeInspectionSetting[] GetDefaultCodeInspections()
        {
            return _inspections.Select(x =>
                        new CodeInspectionSetting(x.Name, x.Description, x.InspectionType, x.DefaultSeverity,
                            x.DefaultSeverity)).ToArray();
        }

        public IndenterSettings GetDefaultIndenterSettings()
        {
            var tabWidth = 4;
            var reg = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\6.0\Common", false) ??
                      Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\VBA\7.0\Common", false);
            if (reg != null)
            {
                tabWidth = Convert.ToInt32(reg.GetValue("TabWidth") ?? tabWidth);
            }
            return new IndenterSettings
            {
                IndentEntireProcedureBody = true,
                IndentFirstCommentBlock = true,
                IndentFirstDeclarationBlock = true,
                AlignCommentsWithCode = true,
                AlignContinuations = true,
                IgnoreOperatorsInContinuations = true,
                IndentCase = false,
                ForceDebugStatementsInColumn1 = false,
                ForceCompilerDirectivesInColumn1 = false,
                IndentCompilerDirectives = true,
                AlignDims = false,
                AlignDimColumn = 15,
                EnableUndo = true,
                EndOfLineCommentStyle = SmartIndenter.EndOfLineCommentStyle.AlignInColumn,
                EndOfLineCommentColumnSpaceAlignment = 50,
                IndentSpaces = tabWidth
            };
        }
    }
}
