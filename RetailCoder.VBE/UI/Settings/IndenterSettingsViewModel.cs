using System;
using System.Linq;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.Settings
{
    public class IndenterSettingsViewModel : SettingsViewModelBase, ISettingsViewModel
    {
        public IndenterSettingsViewModel(Configuration config)
        {
            _alignCommentsWithCode = config.UserSettings.IndenterSettings.AlignCommentsWithCode;
            _alignContinuations = config.UserSettings.IndenterSettings.AlignContinuations;
            _alignDimColumn = config.UserSettings.IndenterSettings.AlignDimColumn;
            _alignDims = config.UserSettings.IndenterSettings.AlignDims;
            _endOfLineCommentColumnSpaceAlignment = config.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment;
            _endOfLineCommentStyle = config.UserSettings.IndenterSettings.EndOfLineCommentStyle;
            _forceCompilerDirectivesInColumn1 = config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1;
            _forceDebugStatementsInColumn1 = config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1;
            _ignoreOperatorsInContinuations = config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations;
            _indentCase = config.UserSettings.IndenterSettings.IndentCase;
            _indentCompilerDirectives = config.UserSettings.IndenterSettings.IndentCompilerDirectives;
            _indentEnumTypeAsProcedure = config.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure;
            _indentEntireProcedureBody = config.UserSettings.IndenterSettings.IndentEntireProcedureBody;
            _indentFirstCommentBlock = config.UserSettings.IndenterSettings.IndentFirstCommentBlock;
            _indentFirstDeclarationBlock = config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock;
            _indentSpaces = config.UserSettings.IndenterSettings.IndentSpaces;
            _spaceProcedures = config.UserSettings.IndenterSettings.VerticallySpaceProcedures;
            _procedureSpacing = config.UserSettings.IndenterSettings.LinesBetweenProcedures;

            PropertyChanged += IndenterSettingsViewModel_PropertyChanged;
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings());
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        void IndenterSettingsViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            // ReSharper disable once ExplicitCallerInfoArgument
            if (e.PropertyName != "PreviewSampleCode")
            {
                OnPropertyChanged("PreviewSampleCode");
            }
        }

        #region Properties

        private bool _alignCommentsWithCode;
        public bool AlignCommentsWithCode
        {
            get { return _alignCommentsWithCode; }
            set 
            { 
                if (_alignCommentsWithCode != value)
                {
                    _alignCommentsWithCode = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _indentEnumTypeAsProcedure;

        public bool IndentEnumTypeAsProcedure
        {
            get { return _indentEnumTypeAsProcedure; }
            set
            {
                if (_indentEnumTypeAsProcedure != value)
                {
                    _indentEnumTypeAsProcedure = value;
                    OnPropertyChanged();
                }
            }           
        }

        private bool _alignContinuations;
        public bool AlignContinuations
        {
            get { return _alignContinuations; }
            set 
            {
                if (_alignContinuations != value)
                {
                    _alignContinuations = value;
                    OnPropertyChanged();
                } 
            }
        }

        private int _alignDimColumn;
        public int AlignDimColumn
        {
            get { return _alignDimColumn; }
            set
            { 
                if (_alignDimColumn != value)
                {
                    _alignDimColumn = value;
                    OnPropertyChanged(); 
                }
            }
        }

        private bool _alignDims;
        public bool AlignDims
        {
            get { return _alignDims; }
            set 
            { 
                if (_alignDims != value)
                {
                    _alignDims = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _endOfLineCommentColumnSpaceAlignment;
        public int EndOfLineCommentColumnSpaceAlignment
        {
            get { return _endOfLineCommentColumnSpaceAlignment; }
            set
            {
                if (_endOfLineCommentColumnSpaceAlignment != value)
                {
                    _endOfLineCommentColumnSpaceAlignment = value;
                    OnPropertyChanged();
                }
            }
        }

        private EndOfLineCommentStyle _endOfLineCommentStyle;
        public EndOfLineCommentStyle EndOfLineCommentStyle
        {
            get { return _endOfLineCommentStyle; }
            set 
            {
                if (_endOfLineCommentStyle != value)
                {
                    _endOfLineCommentStyle = value;
                    OnPropertyChanged(); 
                }
            }
        }

        private bool _forceCompilerDirectivesInColumn1;
        public bool ForceCompilerDirectivesInColumn1
        {
            get { return _forceCompilerDirectivesInColumn1; }
            set 
            {
                if (_forceCompilerDirectivesInColumn1 != value)
                {
                    _forceCompilerDirectivesInColumn1 = value;
                    OnPropertyChanged(); 
                } 
            }
        }

        private bool _forceDebugStatementsInColumn1;
        public bool ForceDebugStatementsInColumn1
        {
            get { return _forceDebugStatementsInColumn1; }
            set 
            {
                if (_forceDebugStatementsInColumn1 != value)
                {
                    _forceDebugStatementsInColumn1 = value;OnPropertyChanged(); 
                }
            }
        }

        private bool _ignoreOperatorsInContinuations;
        public bool IgnoreOperatorsInContinuations
        {
            get { return _ignoreOperatorsInContinuations; }
            set
            {
                if (_ignoreOperatorsInContinuations != value)
                {
                    _ignoreOperatorsInContinuations = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _indentCase;
        public bool IndentCase
        {
            get { return _indentCase; }
            set
            {
                if (_indentCase != value)
                {
                    _indentCase = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _indentCompilerDirectives;
        public bool IndentCompilerDirectives
        {
            get { return _indentCompilerDirectives; }
            set
            {
                if (_indentCompilerDirectives != value)
                {
                    _indentCompilerDirectives = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _indentEntireProcedureBody;
        public bool IndentEntireProcedureBody
        {
            get { return _indentEntireProcedureBody; }
            set
            {
                if (_indentEntireProcedureBody != value)
                {
                    _indentEntireProcedureBody = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _indentFirstCommentBlock;
        public bool IndentFirstCommentBlock
        {
            get { return _indentFirstCommentBlock; }
            set
            {
                if (_indentFirstCommentBlock != value)
                {
                    _indentFirstCommentBlock = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _indentFirstDeclarationBlock;
        public bool IndentFirstDeclarationBlock
        {
            get { return _indentFirstDeclarationBlock; }
            set
            {
                if (_indentFirstDeclarationBlock != value)
                {
                    _indentFirstDeclarationBlock = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _indentSpaces;
        public int IndentSpaces
        {
            get { return _indentSpaces; }
            set
            {
                if (_indentSpaces != value)
                {
                    _indentSpaces = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _spaceProcedures;
        public bool VerticallySpaceProcedures
        {
            get { return _spaceProcedures; }
            set
            {
                if (_spaceProcedures != value)
                {
                    _spaceProcedures = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _procedureSpacing;
        public int LinesBetweenProcedures
        {
            get { return _procedureSpacing; }
            set
            {
                if (_procedureSpacing != value)
                {
                    _procedureSpacing = value;
                    OnPropertyChanged();
                }
            }
        }

        public string PreviewSampleCode 
        {
            get
            {
                var indenter = new Indenter(null, GetCurrentSettings);

                var lines = RubberduckUI.IndenterSettings_PreviewCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                lines = indenter.Indent(lines).ToArray();
                return string.Join(Environment.NewLine, lines);
            }
        }

        private IIndenterSettings GetCurrentSettings()
        {
            return new SmartIndenter.IndenterSettings
            {
                AlignCommentsWithCode = AlignCommentsWithCode,
                AlignContinuations = AlignContinuations,
                AlignDimColumn = AlignDimColumn,
                AlignDims = AlignDims,
                EndOfLineCommentColumnSpaceAlignment = EndOfLineCommentColumnSpaceAlignment,
                EndOfLineCommentStyle = EndOfLineCommentStyle,
                ForceCompilerDirectivesInColumn1 = ForceCompilerDirectivesInColumn1,
                ForceDebugStatementsInColumn1 = ForceDebugStatementsInColumn1,
                IgnoreOperatorsInContinuations = IgnoreOperatorsInContinuations,
                IndentCase = IndentCase,
                IndentCompilerDirectives = IndentCompilerDirectives,
                IndentEnumTypeAsProcedure = IndentEnumTypeAsProcedure,
                IndentEntireProcedureBody = IndentEntireProcedureBody,
                IndentFirstCommentBlock = IndentFirstCommentBlock,
                IndentFirstDeclarationBlock = IndentFirstDeclarationBlock,
                IndentSpaces = IndentSpaces,
                VerticallySpaceProcedures = VerticallySpaceProcedures,
                LinesBetweenProcedures = LinesBetweenProcedures
            };
        }

        #endregion

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.IndenterSettings.AlignCommentsWithCode = AlignCommentsWithCode;
            config.UserSettings.IndenterSettings.AlignContinuations = AlignContinuations;
            config.UserSettings.IndenterSettings.AlignDimColumn = AlignDimColumn;
            config.UserSettings.IndenterSettings.AlignDims = AlignDims;
            config.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment = EndOfLineCommentColumnSpaceAlignment;
            config.UserSettings.IndenterSettings.EndOfLineCommentStyle = EndOfLineCommentStyle;
            config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1 = ForceCompilerDirectivesInColumn1;
            config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1 = ForceDebugStatementsInColumn1;
            config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations = IgnoreOperatorsInContinuations;
            config.UserSettings.IndenterSettings.IndentCase = IndentCase;
            config.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure = IndentEnumTypeAsProcedure;
            config.UserSettings.IndenterSettings.IndentCompilerDirectives = IndentCompilerDirectives;
            config.UserSettings.IndenterSettings.IndentEntireProcedureBody = IndentEntireProcedureBody;
            config.UserSettings.IndenterSettings.IndentFirstCommentBlock = IndentFirstCommentBlock;
            config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock = IndentFirstDeclarationBlock;
            config.UserSettings.IndenterSettings.IndentSpaces = IndentSpaces;
            config.UserSettings.IndenterSettings.VerticallySpaceProcedures = VerticallySpaceProcedures;
            config.UserSettings.IndenterSettings.LinesBetweenProcedures = LinesBetweenProcedures;
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.IndenterSettings);
        }

        private void TransferSettingsToView(IIndenterSettings toLoad)
        {
            AlignCommentsWithCode = toLoad.AlignCommentsWithCode;
            AlignContinuations = toLoad.AlignContinuations;
            AlignDimColumn = toLoad.AlignDimColumn;
            AlignDims = toLoad.AlignDims;
            EndOfLineCommentColumnSpaceAlignment = toLoad.EndOfLineCommentColumnSpaceAlignment;
            EndOfLineCommentStyle = toLoad.EndOfLineCommentStyle;
            ForceCompilerDirectivesInColumn1 = toLoad.ForceCompilerDirectivesInColumn1;
            ForceDebugStatementsInColumn1 = toLoad.ForceDebugStatementsInColumn1;
            IgnoreOperatorsInContinuations = toLoad.IgnoreOperatorsInContinuations;
            IndentCase = toLoad.IndentCase;
            IndentEnumTypeAsProcedure = toLoad.IndentEnumTypeAsProcedure;
            IndentCompilerDirectives = toLoad.IndentCompilerDirectives;
            IndentEntireProcedureBody = toLoad.IndentEntireProcedureBody;
            IndentFirstCommentBlock = toLoad.IndentFirstCommentBlock;
            IndentFirstDeclarationBlock = toLoad.IndentFirstDeclarationBlock;
            IndentSpaces = toLoad.IndentSpaces;
            VerticallySpaceProcedures = toLoad.VerticallySpaceProcedures;
            LinesBetweenProcedures = toLoad.LinesBetweenProcedures;
        }

        private void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_LoadIndenterSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<SmartIndenter.IndenterSettings> { FilePath = dialog.FileName };
                var loaded = service.Load(new SmartIndenter.IndenterSettings());
                TransferSettingsToView(loaded);
            }
        }

        private void ExportSettings()
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = RubberduckUI.DialogMask_XmlFilesOnly,
                Title = RubberduckUI.DialogCaption_SaveIndenterSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var service = new XmlPersistanceService<SmartIndenter.IndenterSettings> {FilePath = dialog.FileName};
                service.Save((SmartIndenter.IndenterSettings)GetCurrentSettings());
            }
        }
    }
}
