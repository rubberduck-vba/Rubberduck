using NLog;
using Rubberduck.Resources;
using Rubberduck.Resources.Settings;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class AutoCompleteSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public AutoCompleteSettingsViewModel(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.AutoCompleteSettings);
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.AutoCompleteSettings);
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.AutoCompleteSettings.IsEnabled = IsEnabled;

            config.UserSettings.AutoCompleteSettings.SelfClosingPairs.IsEnabled = EnableSelfClosingPairs;

            config.UserSettings.AutoCompleteSettings.SmartConcat.IsEnabled = EnableSmartConcat;
            config.UserSettings.AutoCompleteSettings.SmartConcat.ConcatVbNewLineModifier =
                ConcatVbNewLine ? ModifierKeySetting.CtrlKey : ModifierKeySetting.None;
            config.UserSettings.AutoCompleteSettings.SmartConcat.ConcatMaxLines = ConcatMaxLines;
            config.UserSettings.AutoCompleteSettings.BlockCompletion.IsEnabled = EnableBlockCompletion;
            config.UserSettings.AutoCompleteSettings.BlockCompletion.CompleteOnTab = CompleteBlockOnTab;
            config.UserSettings.AutoCompleteSettings.BlockCompletion.CompleteOnEnter = CompleteBlockOnEnter;
        }

        private void TransferSettingsToView(Rubberduck.Settings.AutoCompleteSettings toLoad)
        {
            IsEnabled = toLoad.IsEnabled;

            EnableSelfClosingPairs = toLoad.SelfClosingPairs.IsEnabled;

            EnableSmartConcat = toLoad.SmartConcat.IsEnabled;
            ConcatVbNewLine = toLoad.SmartConcat.ConcatVbNewLineModifier == ModifierKeySetting.CtrlKey;
            ConcatMaxLines = toLoad.SmartConcat.ConcatMaxLines;

            EnableBlockCompletion = toLoad.BlockCompletion.IsEnabled;
            CompleteBlockOnTab = toLoad.BlockCompletion.CompleteOnTab;
            CompleteBlockOnEnter = toLoad.BlockCompletion.CompleteOnEnter;
        }

        private bool _isEnabled;

        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                if (_isEnabled != value)
                {
                    _isEnabled = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _enableSelfClosingPairs;

        public bool EnableSelfClosingPairs
        {
            get { return _enableSelfClosingPairs; }
            set
            {
                if (_enableSelfClosingPairs != value)
                {
                    _enableSelfClosingPairs = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _enableSmartConcat;
        public bool EnableSmartConcat
        {
            get { return _enableSmartConcat; }
            set
            {
                if (_enableSmartConcat != value)
                {
                    _enableSmartConcat = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _concatVbNewLine;

        public bool ConcatVbNewLine
        {
            get { return _concatVbNewLine; }
            set
            {
                if (_concatVbNewLine != value)
                {
                    _concatVbNewLine = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _concatMaxLines;
        public int ConcatMaxLines
        {
            get { return _concatMaxLines; }
            set
            {
                if (_concatMaxLines != value)
                {
                    _concatMaxLines = value;
                    OnPropertyChanged();
                }
            }
        }

        public int ConcatMaxLinesMinValue => Rubberduck.Settings.AutoCompleteSettings.ConcatMaxLinesMinValue;
        public int ConcatMaxLinesMaxValue => Rubberduck.Settings.AutoCompleteSettings.ConcatMaxLinesMaxValue;

        private bool _enableBlockCompletion;

        public bool EnableBlockCompletion
        {
            get { return _enableBlockCompletion; }
            set
            {
                if (_enableBlockCompletion != value)
                {
                    _enableBlockCompletion = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _completeBlockOnEnter;
        public bool CompleteBlockOnEnter
        {
            get { return _completeBlockOnEnter; }
            set
            {
                if (_completeBlockOnEnter != value)
                {
                    _completeBlockOnEnter = value;
                    OnPropertyChanged();
                    if (!_completeBlockOnTab && !_completeBlockOnEnter)
                    {
                        // one must be enabled...
                        CompleteBlockOnTab = true;
                    }
                }
            }
        }

        private bool _completeBlockOnTab;
        public bool CompleteBlockOnTab
        {
            get { return _completeBlockOnTab; }
            set
            {
                if (_completeBlockOnTab != value)
                {
                    _completeBlockOnTab = value;
                    OnPropertyChanged();
                    if (!_completeBlockOnTab && !_completeBlockOnEnter)
                    {
                        // one must be enabled...
                        CompleteBlockOnEnter = true;
                    }
                }
            }
        }

        private void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = SettingsUI.DialogMask_XmlFilesOnly,
                Title = SettingsUI.DialogCaption_LoadInspectionSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.AutoCompleteSettings> { FilePath = dialog.FileName };
                var loaded = service.Load(new Rubberduck.Settings.AutoCompleteSettings());
                TransferSettingsToView(loaded);
            }
        }

        private void ExportSettings()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = SettingsUI.DialogMask_XmlFilesOnly,
                Title = SettingsUI.DialogCaption_SaveAutocompletionSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<Rubberduck.Settings.AutoCompleteSettings> { FilePath = dialog.FileName };
                service.Save(new Rubberduck.Settings.AutoCompleteSettings
                {
                    IsEnabled = IsEnabled,
                    BlockCompletion = new Rubberduck.Settings.AutoCompleteSettings.BlockCompletionSettings
                    {
                        CompleteOnEnter = CompleteBlockOnEnter,
                        CompleteOnTab = CompleteBlockOnTab,
                        IsEnabled = EnableBlockCompletion
                    },
                    SelfClosingPairs = new Rubberduck.Settings.AutoCompleteSettings.SelfClosingPairSettings
                    {
                        IsEnabled = EnableSelfClosingPairs
                    },
                    SmartConcat = new Rubberduck.Settings.AutoCompleteSettings.SmartConcatSettings
                    {
                        ConcatVbNewLineModifier =
                            ConcatVbNewLine ? ModifierKeySetting.CtrlKey : ModifierKeySetting.None,
                        IsEnabled = EnableSmartConcat
                    }
                });
            }
        }
    }
}
