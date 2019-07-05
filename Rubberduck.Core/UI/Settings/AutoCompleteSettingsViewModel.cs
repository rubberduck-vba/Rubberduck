using System.Windows.Input;
using NLog;
using Rubberduck.Resources.Settings;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public sealed class AutoCompleteSettingsViewModel : SettingsViewModelBase<Rubberduck.Settings.AutoCompleteSettings>, ISettingsViewModel<Rubberduck.Settings.AutoCompleteSettings>
    {
        public AutoCompleteSettingsViewModel(Configuration config, IConfigurationService<Rubberduck.Settings.AutoCompleteSettings> service) 
            : base(service)
        {
            TransferSettingsToView(config.UserSettings.AutoCompleteSettings);
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ =>
                ExportSettings(new Rubberduck.Settings.AutoCompleteSettings
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
                }));
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());

            IncrementMaxConcatLinesCommand = new DelegateCommand(null, ExecuteIncrementMaxConcatLines, CanExecuteIncrementMaxConcatLines);
            DecrementMaxConcatLinesCommand = new DelegateCommand(null, ExecuteDecrementMaxConcatLines, CanExecuteDecrementMaxConcatLines);
        }

        public ICommand IncrementMaxConcatLinesCommand { get; }

        private bool CanExecuteIncrementMaxConcatLines(object parameter) => ConcatMaxLines < ConcatMaxLinesMaxValue;
        private void ExecuteIncrementMaxConcatLines(object parameter) => ConcatMaxLines++;

        public ICommand DecrementMaxConcatLinesCommand { get; }

        private bool CanExecuteDecrementMaxConcatLines(object parameter) => ConcatMaxLines > ConcatMaxLinesMinValue;
        private void ExecuteDecrementMaxConcatLines(object parameter) => ConcatMaxLines--;

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

        protected override void TransferSettingsToView(Rubberduck.Settings.AutoCompleteSettings toLoad)
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

        protected override string DialogLoadTitle => SettingsUI.DialogCaption_LoadInspectionSettings;
        protected override string DialogSaveTitle => SettingsUI.DialogCaption_SaveAutocompletionSettings;
    }
}
