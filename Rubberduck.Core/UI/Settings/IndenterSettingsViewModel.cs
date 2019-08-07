using System;
using System.Linq;
using NLog;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command;
using Rubberduck.Resources;
using Rubberduck.Resources.Settings;

namespace Rubberduck.UI.Settings
{
    public sealed class IndenterSettingsViewModel : SettingsViewModelBase<SmartIndenter.IndenterSettings>, ISettingsViewModel<SmartIndenter.IndenterSettings>
    {
        public IndenterSettingsViewModel(Configuration config, IConfigurationService<SmartIndenter.IndenterSettings> service)
            : base(service)
        {
            _alignCommentsWithCode = config.UserSettings.IndenterSettings.AlignCommentsWithCode;
            _alignContinuations = config.UserSettings.IndenterSettings.AlignContinuations;
            _alignDimColumn = config.UserSettings.IndenterSettings.AlignDimColumn;
            _alignDims = config.UserSettings.IndenterSettings.AlignDims;
            _endOfLineCommentColumnSpaceAlignment = config.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment;
            _endOfLineCommentStyle = config.UserSettings.IndenterSettings.EndOfLineCommentStyle;
            _forceCompilerDirectivesInColumn1 = config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1;
            _forceDebugStatementsInColumn1 = config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1;
            _forceDebugPrintInColumn1 = config.UserSettings.IndenterSettings.ForceDebugPrintInColumn1;
            _forceDebugAssertInColumn1 = config.UserSettings.IndenterSettings.ForceDebugAssertInColumn1;
            _forceStopInColumn1 = config.UserSettings.IndenterSettings.ForceStopInColumn1;
            _ignoreOperatorsInContinuations = config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations;
            _indentCase = config.UserSettings.IndenterSettings.IndentCase;
            _indentCompilerDirectives = config.UserSettings.IndenterSettings.IndentCompilerDirectives;
            _indentEnumTypeAsProcedure = config.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure;
            _indentEntireProcedureBody = config.UserSettings.IndenterSettings.IndentEntireProcedureBody;
            _indentFirstCommentBlock = config.UserSettings.IndenterSettings.IndentFirstCommentBlock;
            _indentFirstDeclarationBlock = config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock;
            _ignoreEmptyLinesInFirstBlocks = config.UserSettings.IndenterSettings.IgnoreEmptyLinesInFirstBlocks;
            _indentSpaces = config.UserSettings.IndenterSettings.IndentSpaces;
            _spaceProcedures = config.UserSettings.IndenterSettings.VerticallySpaceProcedures;
            _procedureSpacing = config.UserSettings.IndenterSettings.LinesBetweenProcedures;

            PropertyChanged += IndenterSettingsViewModel_PropertyChanged;
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings(GetCurrentSettings()));
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());
        }

        protected override string DialogLoadTitle => SettingsUI.DialogCaption_LoadIndenterSettings;
        protected override string DialogSaveTitle => SettingsUI.DialogCaption_SaveIndenterSettings;

        private void IndenterSettingsViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            // ReSharper disable once ExplicitCallerInfoArgument
            if (e.PropertyName != nameof(PreviewSampleCode))
            {
                OnPropertyChanged(nameof(PreviewSampleCode));
            }
        }

        #region Properties

        private bool _alignCommentsWithCode;
        public bool AlignCommentsWithCode
        {
            get => _alignCommentsWithCode;
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
            get => _indentEnumTypeAsProcedure;
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
            get => _alignContinuations;
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
            get => _alignDimColumn;
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
            get => _alignDims;
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
            get => _endOfLineCommentColumnSpaceAlignment;
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
            get => _endOfLineCommentStyle;
            set 
            {
                if (_endOfLineCommentStyle != value)
                {
                    _endOfLineCommentStyle = value;
                    OnPropertyChanged(); 
                }
            }
        }

        private EmptyLineHandling _emptyLineHandlingMethod;
        public EmptyLineHandling EmptyLineHandlingMethod
        {
            get => _emptyLineHandlingMethod;
            set
            {
                if (_emptyLineHandlingMethod != value)
                {
                    _emptyLineHandlingMethod = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _forceCompilerDirectivesInColumn1;
        public bool ForceCompilerDirectivesInColumn1
        {
            get => _forceCompilerDirectivesInColumn1;
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
            get => _forceDebugStatementsInColumn1;
            set 
            {
                if (_forceDebugStatementsInColumn1 != value)
                {
                    _forceDebugStatementsInColumn1 = value;
                    ForceDebugPrintInColumn1 = _forceDebugStatementsInColumn1;
                    ForceDebugAssertInColumn1 = _forceDebugStatementsInColumn1;
                    ForceStopInColumn1 = _forceDebugStatementsInColumn1;
                    OnPropertyChanged(); 
                }
            }
        }

        private bool _forceDebugPrintInColumn1;
        public bool ForceDebugPrintInColumn1
        {
            get => _forceDebugPrintInColumn1;
            set
            {
                if (_forceDebugPrintInColumn1 != value)
                {
                    _forceDebugPrintInColumn1 = value;
                    if (!_forceDebugPrintInColumn1 && !_forceDebugAssertInColumn1 && !_forceStopInColumn1)
                    {
                        ForceDebugStatementsInColumn1 = false;
                    }
                    OnPropertyChanged();
                }
            }
        }

        private bool _forceDebugAssertInColumn1;
        public bool ForceDebugAssertInColumn1
        {
            get => _forceDebugAssertInColumn1;
            set
            {
                if (_forceDebugAssertInColumn1 != value)
                {
                    _forceDebugAssertInColumn1 = value;
                    if (!_forceDebugPrintInColumn1 && !_forceDebugAssertInColumn1 && !_forceStopInColumn1)
                    {
                        ForceDebugStatementsInColumn1 = false;
                    }
                    OnPropertyChanged();
                }
            }
        }

        private bool _forceStopInColumn1;
        public bool ForceStopInColumn1
        {
            get => _forceStopInColumn1;
            set
            {
                if (_forceStopInColumn1 != value)
                {
                    _forceStopInColumn1 = value;
                    if (!_forceDebugPrintInColumn1 && !_forceDebugAssertInColumn1 && !_forceStopInColumn1)
                    {
                        ForceDebugStatementsInColumn1 = false;
                    }
                    OnPropertyChanged();
                }
            }
        }    

        private bool _ignoreOperatorsInContinuations;
        public bool IgnoreOperatorsInContinuations
        {
            get => _ignoreOperatorsInContinuations;
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
            get => _indentCase;
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
            get => _indentCompilerDirectives;
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
            get => _indentEntireProcedureBody;
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
            get => _indentFirstCommentBlock;
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
            get => _indentFirstDeclarationBlock;
            set
            {
                if (_indentFirstDeclarationBlock != value)
                {
                    _indentFirstDeclarationBlock = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _ignoreEmptyLinesInFirstBlocks;
        public bool IgnoreEmptyLinesInFirstBlocks
        {
            get => _ignoreEmptyLinesInFirstBlocks;
            set
            {
                if (_ignoreEmptyLinesInFirstBlocks != value)
                {
                    _ignoreEmptyLinesInFirstBlocks = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _indentSpaces;
        public int IndentSpaces
        {
            get => _indentSpaces;
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
            get => _spaceProcedures;
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
            get => _procedureSpacing;
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

        private SmartIndenter.IndenterSettings GetCurrentSettings()
        {
            return new SmartIndenter.IndenterSettings(false)
            {
                AlignCommentsWithCode = AlignCommentsWithCode,
                AlignContinuations = AlignContinuations,
                AlignDimColumn = AlignDimColumn,
                AlignDims = AlignDims,
                EndOfLineCommentColumnSpaceAlignment = EndOfLineCommentColumnSpaceAlignment,
                EndOfLineCommentStyle = EndOfLineCommentStyle,
                EmptyLineHandlingMethod = EmptyLineHandlingMethod,
                ForceCompilerDirectivesInColumn1 = ForceCompilerDirectivesInColumn1,
                ForceDebugStatementsInColumn1 = ForceDebugStatementsInColumn1,
                ForceDebugPrintInColumn1 = ForceDebugPrintInColumn1,
                ForceDebugAssertInColumn1 = ForceDebugAssertInColumn1,
                ForceStopInColumn1 = ForceStopInColumn1,
                IgnoreOperatorsInContinuations = IgnoreOperatorsInContinuations,
                IndentCase = IndentCase,
                IndentCompilerDirectives = IndentCompilerDirectives,
                IndentEnumTypeAsProcedure = IndentEnumTypeAsProcedure,
                IndentEntireProcedureBody = IndentEntireProcedureBody,
                IndentFirstCommentBlock = IndentFirstCommentBlock,
                IndentFirstDeclarationBlock = IndentFirstDeclarationBlock,
                IgnoreEmptyLinesInFirstBlocks = IgnoreEmptyLinesInFirstBlocks,
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
            config.UserSettings.IndenterSettings.EmptyLineHandlingMethod = EmptyLineHandlingMethod;
            config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1 = ForceCompilerDirectivesInColumn1;
            config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1 = ForceDebugStatementsInColumn1;
            config.UserSettings.IndenterSettings.ForceDebugPrintInColumn1 = ForceDebugPrintInColumn1;
            config.UserSettings.IndenterSettings.ForceDebugAssertInColumn1 = ForceDebugAssertInColumn1;
            config.UserSettings.IndenterSettings.ForceStopInColumn1 = ForceStopInColumn1;
            config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations = IgnoreOperatorsInContinuations;
            config.UserSettings.IndenterSettings.IndentCase = IndentCase;
            config.UserSettings.IndenterSettings.IndentEnumTypeAsProcedure = IndentEnumTypeAsProcedure;
            config.UserSettings.IndenterSettings.IndentCompilerDirectives = IndentCompilerDirectives;
            config.UserSettings.IndenterSettings.IndentEntireProcedureBody = IndentEntireProcedureBody;
            config.UserSettings.IndenterSettings.IndentFirstCommentBlock = IndentFirstCommentBlock;            
            config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock = IndentFirstDeclarationBlock;
            config.UserSettings.IndenterSettings.IgnoreEmptyLinesInFirstBlocks = IgnoreEmptyLinesInFirstBlocks;
            config.UserSettings.IndenterSettings.IndentSpaces = IndentSpaces;
            config.UserSettings.IndenterSettings.VerticallySpaceProcedures = VerticallySpaceProcedures;
            config.UserSettings.IndenterSettings.LinesBetweenProcedures = LinesBetweenProcedures;
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.IndenterSettings);
        }

        protected override void TransferSettingsToView(SmartIndenter.IndenterSettings toLoad)
        {
            AlignCommentsWithCode = toLoad.AlignCommentsWithCode;
            AlignContinuations = toLoad.AlignContinuations;
            AlignDimColumn = toLoad.AlignDimColumn;
            AlignDims = toLoad.AlignDims;
            EndOfLineCommentColumnSpaceAlignment = toLoad.EndOfLineCommentColumnSpaceAlignment;
            EndOfLineCommentStyle = toLoad.EndOfLineCommentStyle;
            EmptyLineHandlingMethod = toLoad.EmptyLineHandlingMethod;
            ForceCompilerDirectivesInColumn1 = toLoad.ForceCompilerDirectivesInColumn1;
            ForceDebugStatementsInColumn1 = toLoad.ForceDebugStatementsInColumn1;
            ForceDebugPrintInColumn1 = toLoad.ForceDebugPrintInColumn1;
            ForceDebugAssertInColumn1 = toLoad.ForceDebugAssertInColumn1;
            ForceStopInColumn1 = toLoad.ForceStopInColumn1;
            IgnoreOperatorsInContinuations = toLoad.IgnoreOperatorsInContinuations;
            IndentCase = toLoad.IndentCase;
            IndentEnumTypeAsProcedure = toLoad.IndentEnumTypeAsProcedure;
            IndentCompilerDirectives = toLoad.IndentCompilerDirectives;
            IndentEntireProcedureBody = toLoad.IndentEntireProcedureBody;
            IndentFirstCommentBlock = toLoad.IndentFirstCommentBlock;
            IndentFirstDeclarationBlock = toLoad.IndentFirstDeclarationBlock;
            IgnoreEmptyLinesInFirstBlocks = toLoad.IgnoreEmptyLinesInFirstBlocks;
            IndentSpaces = toLoad.IndentSpaces;
            VerticallySpaceProcedures = toLoad.VerticallySpaceProcedures;
            LinesBetweenProcedures = toLoad.LinesBetweenProcedures;
        }
    }
}
