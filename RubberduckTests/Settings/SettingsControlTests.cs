using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;
using GeneralSettings = Rubberduck.UI.Settings.GeneralSettings;
using IndenterSettings = Rubberduck.UI.Settings.IndenterSettings;
using UnitTestSettings = Rubberduck.UI.Settings.UnitTestSettings;

namespace RubberduckTests.Settings
{
    [TestClass]
    [Ignore] // these tests incur IO and actually modify the config file.
    public class SettingsControlTests
    {
        private Configuration GetDefaultConfig()
        {
            var generalSettings = new Rubberduck.Settings.GeneralSettings
            {
                Language = new DisplayLanguageSetting("en-US"),
                HotkeySettings = new[]
                {
                    new HotkeySetting{Name="IndentProcedure", IsEnabled=true, Key1="^P"},
                    new HotkeySetting{Name="IndentModule", IsEnabled=true, Key1="^M"}
                },
                AutoSaveEnabled = false,
                AutoSavePeriod = 10
            };

            var todoSettings = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("NOTE "),
                    new ToDoMarker("TODO "),
                    new ToDoMarker("BUG ")
                }
            };

            var inspectionSettings = new CodeInspectionSettings
            {
                CodeInspections = Inspections().Select(i => new CodeInspectionSetting(i)).ToArray()
            };

            var unitTestSettings = new Rubberduck.Settings.UnitTestSettings
            {
                BindingMode = BindingMode.LateBinding,
                AssertMode = AssertMode.StrictAssert,
                ModuleInit = true,
                MethodInit = true,
                DefaultTestStubInNewModule = false
            };

            var indenterSettings = new Rubberduck.Settings.IndenterSettings
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
                EndOfLineCommentStyle = Rubberduck.SmartIndenter.EndOfLineCommentStyle.AlignInColumn,
                EndOfLineCommentColumnSpaceAlignment = 50,
                IndentSpaces = 4
            };

            var userSettings = new UserSettings(generalSettings, todoSettings, inspectionSettings, unitTestSettings, indenterSettings);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var generalSettings = new Rubberduck.Settings.GeneralSettings
            {
                Language = new DisplayLanguageSetting("sv-SE"),
                HotkeySettings = new[]
                {
                    new HotkeySetting{Name="IndentProcedure", IsEnabled=false, Key1="CTRL-C"},
                    new HotkeySetting{Name="IndentModule", IsEnabled=false, Key1="CTRL-X"}
                },
                AutoSaveEnabled = true,
                AutoSavePeriod = 5
            };

            var todoSettings = new ToDoListSettings
            {
                ToDoMarkers = new[]
                {
                    new ToDoMarker("PLACEHOLDER ")
                }
            };

            var inspections = Inspections().Select(i => new CodeInspectionSetting(i)).ToArray();
            inspections[0].Severity = CodeInspectionSeverity.Warning;
            inspections[1].Severity = CodeInspectionSeverity.Suggestion;
            inspections[2].Severity = CodeInspectionSeverity.Hint;
            inspections[3].Severity = CodeInspectionSeverity.Error;
            inspections[4].Severity = CodeInspectionSeverity.DoNotShow;
            inspections[5].Severity = CodeInspectionSeverity.Error;

            var inspectionSettings = new CodeInspectionSettings
            {
                CodeInspections = inspections
            };

            var unitTestSettings = new Rubberduck.Settings.UnitTestSettings
            {
                BindingMode = BindingMode.EarlyBinding,
                AssertMode = AssertMode.PermissiveAssert,
                ModuleInit = false,
                MethodInit = false,
                DefaultTestStubInNewModule = true
            };

            var indenterSettings = new Rubberduck.Settings.IndenterSettings
            {
                IndentEntireProcedureBody = false,
                IndentFirstCommentBlock = false,
                IndentFirstDeclarationBlock = false,
                AlignCommentsWithCode = false,
                AlignContinuations = false,
                IgnoreOperatorsInContinuations = false,
                IndentCase = true,
                ForceDebugStatementsInColumn1 = true,
                ForceCompilerDirectivesInColumn1 = true,
                IndentCompilerDirectives = false,
                AlignDims = true,
                AlignDimColumn = 16,
                EnableUndo = false,
                EndOfLineCommentStyle = Rubberduck.SmartIndenter.EndOfLineCommentStyle.Absolute,
                EndOfLineCommentColumnSpaceAlignment = 60,
                IndentSpaces = 2
            };

            var userSettings = new UserSettings(generalSettings, todoSettings, inspectionSettings, unitTestSettings, indenterSettings);
            return new Configuration(userSettings);
        }

        private IEnumerable<IInspection> Inspections()
        {
            return new IInspection[]
            {
                new AssignedByValParameterInspection(null),
                new ConstantNotUsedInspection(null),
                new DefaultProjectNameInspection(null),
                new EmptyStringLiteralInspection(null),
                new EncapsulatePublicFieldInspection(null),
                new MoveFieldCloserToUsageInspection(null),
                new NonReturningFunctionInspection(null),
                new ObsoleteCallStatementInspection(null),
                new ProcedureNotUsedInspection(null)
            };
        }

        private ConfigurationLoader GetConfigLoader(Configuration config)
        {
            var configLoader = new ConfigurationLoader(Inspections());
            configLoader.SaveConfiguration(config);

            return configLoader;
        }

        private SettingsControlViewModel GetDefaultViewModel(ConfigurationLoader configService, SettingsViews activeView = SettingsViews.GeneralSettings)
        {
            var config = configService.LoadConfiguration();

            return new SettingsControlViewModel(configService,
                config,
                new SettingsView
                {
                    Control = new GeneralSettings(new GeneralSettingsViewModel(config)),
                    View = SettingsViews.GeneralSettings
                },
                new SettingsView
                {
                    Control = new TodoSettings(new TodoSettingsViewModel(config)),
                    View = SettingsViews.TodoSettings
                },
                new SettingsView
                {
                    Control = new InspectionSettings(new InspectionSettingsViewModel(config)),
                    View = SettingsViews.InspectionSettings
                },
                new SettingsView
                {
                    Control = new UnitTestSettings(new UnitTestSettingsViewModel(config)),
                    View = SettingsViews.UnitTestSettings
                },
                new SettingsView
                {
                    Control = new IndenterSettings(new IndenterSettingsViewModel(config)),
                    View = SettingsViews.IndenterSettings
                },
                activeView);
        }

        [TestMethod]
        public void DefaultViewIsGeneralSettings()
        {
            var viewModel = GetDefaultViewModel(GetConfigLoader(GetDefaultConfig()));

            Assert.AreEqual(SettingsViews.GeneralSettings, viewModel.SelectedSettingsView.View);
        }

        [TestMethod]
        public void PassedInViewIsSelected()
        {
            var viewModel = GetDefaultViewModel(GetConfigLoader(GetDefaultConfig()), SettingsViews.TodoSettings);

            Assert.AreEqual(SettingsViews.TodoSettings, viewModel.SelectedSettingsView.View);
        }

        [TestMethod]
        public void OKButtonSavesConfigSave()
        {
            var nondefaultConfig = GetNondefaultConfig();
            var configLoader = GetConfigLoader(nondefaultConfig);
            var viewModel = GetDefaultViewModel(configLoader);

            var defaultConfig = GetDefaultConfig();
            foreach (var view in viewModel.SettingsViews)
            {
                view.Control.ViewModel.SetToDefaults(defaultConfig);
            }

            viewModel.OKButtonCommand.Execute(null);

            var updatedConfig = configLoader.LoadConfiguration();

            MultiAssert.Aggregate(
                () => MultiAssert.Aggregate(
                          () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, updatedConfig.UserSettings.GeneralSettings.Language),
                          () => Assert.IsTrue(defaultConfig.UserSettings.GeneralSettings.HotkeySettings.SequenceEqual(updatedConfig.UserSettings.GeneralSettings.HotkeySettings)),
                          () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSaveEnabled, updatedConfig.UserSettings.GeneralSettings.AutoSaveEnabled),
                          () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, updatedConfig.UserSettings.GeneralSettings.AutoSavePeriod)
                      ),
                () => Assert.IsTrue(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.SequenceEqual(updatedConfig.UserSettings.ToDoListSettings.ToDoMarkers)),
                () => Assert.IsTrue(defaultConfig.UserSettings.CodeInspectionSettings.CodeInspections.SequenceEqual(updatedConfig.UserSettings.CodeInspectionSettings.CodeInspections)),
                () => MultiAssert.Aggregate(
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.BindingMode, updatedConfig.UserSettings.UnitTestSettings.BindingMode),
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.AssertMode, updatedConfig.UserSettings.UnitTestSettings.AssertMode),
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.ModuleInit, updatedConfig.UserSettings.UnitTestSettings.ModuleInit),
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.MethodInit, updatedConfig.UserSettings.UnitTestSettings.MethodInit),
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule, updatedConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule)
                      ),
                () => MultiAssert.Aggregate(
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignCommentsWithCode, updatedConfig.UserSettings.IndenterSettings.AlignCommentsWithCode),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignContinuations, updatedConfig.UserSettings.IndenterSettings.AlignContinuations),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDimColumn, updatedConfig.UserSettings.IndenterSettings.AlignDimColumn),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDims, updatedConfig.UserSettings.IndenterSettings.AlignDims),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EnableUndo, updatedConfig.UserSettings.IndenterSettings.EnableUndo),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, updatedConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle, updatedConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, updatedConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, updatedConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, updatedConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCase, updatedConfig.UserSettings.IndenterSettings.IndentCase),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCompilerDirectives, updatedConfig.UserSettings.IndenterSettings.IndentCompilerDirectives),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody, updatedConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock, updatedConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, updatedConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentSpaces, updatedConfig.UserSettings.IndenterSettings.IndentSpaces)));
        }

        [TestMethod]
        public void CancelButtonDoesNothing()
        {
            var nondefaultConfig = GetNondefaultConfig();
            var configLoader = GetConfigLoader(nondefaultConfig);
            var viewModel = GetDefaultViewModel(configLoader);

            var defaultConfig = GetDefaultConfig();
            foreach (var view in viewModel.SettingsViews)
            {
                view.Control.ViewModel.SetToDefaults(defaultConfig);
            }

            viewModel.CancelButtonCommand.Execute(null);

            var updatedConfig = configLoader.LoadConfiguration();

            MultiAssert.Aggregate(
                () => MultiAssert.Aggregate(
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.GeneralSettings.Language, updatedConfig.UserSettings.GeneralSettings.Language),
                          () => Assert.IsTrue(nondefaultConfig.UserSettings.GeneralSettings.HotkeySettings.SequenceEqual(updatedConfig.UserSettings.GeneralSettings.HotkeySettings)),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.GeneralSettings.AutoSaveEnabled, updatedConfig.UserSettings.GeneralSettings.AutoSaveEnabled),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, updatedConfig.UserSettings.GeneralSettings.AutoSavePeriod)
                      ),
                () => Assert.IsTrue(nondefaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.SequenceEqual(updatedConfig.UserSettings.ToDoListSettings.ToDoMarkers)),
                () => Assert.IsTrue(nondefaultConfig.UserSettings.CodeInspectionSettings.CodeInspections.SequenceEqual(updatedConfig.UserSettings.CodeInspectionSettings.CodeInspections)),
                () => MultiAssert.Aggregate(
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.UnitTestSettings.BindingMode, updatedConfig.UserSettings.UnitTestSettings.BindingMode),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.UnitTestSettings.AssertMode, updatedConfig.UserSettings.UnitTestSettings.AssertMode),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.UnitTestSettings.ModuleInit, updatedConfig.UserSettings.UnitTestSettings.ModuleInit),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.UnitTestSettings.MethodInit, updatedConfig.UserSettings.UnitTestSettings.MethodInit),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule, updatedConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule)
                      ),
                () => MultiAssert.Aggregate(
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.AlignCommentsWithCode, updatedConfig.UserSettings.IndenterSettings.AlignCommentsWithCode),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.AlignContinuations, updatedConfig.UserSettings.IndenterSettings.AlignContinuations),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.AlignDimColumn, updatedConfig.UserSettings.IndenterSettings.AlignDimColumn),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.AlignDims, updatedConfig.UserSettings.IndenterSettings.AlignDims),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.EnableUndo, updatedConfig.UserSettings.IndenterSettings.EnableUndo),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, updatedConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle, updatedConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, updatedConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, updatedConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, updatedConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.IndentCase, updatedConfig.UserSettings.IndenterSettings.IndentCase),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.IndentCompilerDirectives, updatedConfig.UserSettings.IndenterSettings.IndentCompilerDirectives),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody, updatedConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock, updatedConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, updatedConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock),
                          () => Assert.AreEqual(nondefaultConfig.UserSettings.IndenterSettings.IndentSpaces, updatedConfig.UserSettings.IndenterSettings.IndentSpaces)));
        }

        [TestMethod]
        public void ResetButtonResetsVMs()
        {
            var nondefaultConfig = GetNondefaultConfig();
            var configLoader = GetConfigLoader(nondefaultConfig);
            var viewModel = GetDefaultViewModel(configLoader);

            var defaultConfig = GetDefaultConfig();

            viewModel.ResetButtonCommand.Execute(null);

            MultiAssert.Aggregate(
                () => MultiAssert.Aggregate(
                          () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.Language, ((GeneralSettingsViewModel)viewModel.SettingsViews[0].Control.ViewModel).SelectedLanguage),
                          () => Assert.IsTrue(defaultConfig.UserSettings.GeneralSettings.HotkeySettings.SequenceEqual(((GeneralSettingsViewModel)viewModel.SettingsViews[0].Control.ViewModel).Hotkeys)),
                          () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSaveEnabled, ((GeneralSettingsViewModel)viewModel.SettingsViews[0].Control.ViewModel).AutoSaveEnabled),
                          () => Assert.AreEqual(defaultConfig.UserSettings.GeneralSettings.AutoSavePeriod, ((GeneralSettingsViewModel)viewModel.SettingsViews[0].Control.ViewModel).AutoSavePeriod)
                      ),
                () => Assert.IsTrue(defaultConfig.UserSettings.ToDoListSettings.ToDoMarkers.SequenceEqual(viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<TodoSettingsViewModel>().First().TodoSettings)),
                () => Assert.IsTrue(defaultConfig.UserSettings.CodeInspectionSettings.CodeInspections.SequenceEqual(viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<InspectionSettingsViewModel>().First().InspectionSettings.SourceCollection.Cast<CodeInspectionSetting>()), "test"),
                () => MultiAssert.Aggregate(
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.BindingMode, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<UnitTestSettingsViewModel>().First().BindingMode),
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.AssertMode, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<UnitTestSettingsViewModel>().First().AssertMode),
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.ModuleInit, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<UnitTestSettingsViewModel>().First().ModuleInit),
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.MethodInit, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<UnitTestSettingsViewModel>().First().MethodInit),
                          () => Assert.AreEqual(defaultConfig.UserSettings.UnitTestSettings.DefaultTestStubInNewModule, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<UnitTestSettingsViewModel>().First().DefaultTestStubInNewModule)
                      ),
                () => MultiAssert.Aggregate(
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().AlignCommentsWithCode),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignContinuations, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().AlignContinuations),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDimColumn, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().AlignDimColumn),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDims, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().AlignDims),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EnableUndo, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().EnableUndo),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().EndOfLineCommentColumnSpaceAlignment),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().EndOfLineCommentStyle),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().ForceCompilerDirectivesInColumn1),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().ForceDebugStatementsInColumn1),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().IgnoreOperatorsInContinuations),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCase, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().IndentCase),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().IndentCompilerDirectives),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().IndentEntireProcedureBody),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().IndentFirstCommentBlock),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().IndentFirstDeclarationBlock),
                          () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentSpaces, viewModel.SettingsViews.Select(v => v.Control.ViewModel).OfType<IndenterSettingsViewModel>().First().IndentSpaces)));
        }

        [TestMethod]
        public void OKButton_ReusesAction()
        {
            var viewModel = GetDefaultViewModel(GetConfigLoader(GetDefaultConfig()));

            var initialOKButtonCommand = viewModel.OKButtonCommand;
            Assert.AreSame(initialOKButtonCommand, viewModel.OKButtonCommand);
        }

        [TestMethod]
        public void CancelButton_ReusesAction()
        {
            var viewModel = GetDefaultViewModel(GetConfigLoader(GetDefaultConfig()));

            var initialCancelButtonCommand = viewModel.CancelButtonCommand;
            Assert.AreSame(initialCancelButtonCommand, viewModel.CancelButtonCommand);
        }

        [TestMethod]
        public void ResetButton_ReusesAction()
        {
            var viewModel = GetDefaultViewModel(GetConfigLoader(GetDefaultConfig()));

            var initialResetButtonCommand = viewModel.ResetButtonCommand;
            Assert.AreSame(initialResetButtonCommand, viewModel.ResetButtonCommand);
        }

        [TestMethod]
        public void OKButtonFiresEvent()
        {
            var eventIsFired = false;
            var viewModel = GetDefaultViewModel(GetConfigLoader(GetDefaultConfig()));

            viewModel.OnWindowClosed += (sender, args) => { eventIsFired = true; };

            viewModel.OKButtonCommand.Execute(null);

            Assert.IsTrue(eventIsFired);
        }

        [TestMethod]
        public void CancelButtonFiresEvent()
        {
            var eventIsFired = false;
            var viewModel = GetDefaultViewModel(GetConfigLoader(GetDefaultConfig()));

            viewModel.OnWindowClosed += (sender, args) => { eventIsFired = true; };

            viewModel.CancelButtonCommand.Execute(null);

            Assert.IsTrue(eventIsFired);
        }
    }
}