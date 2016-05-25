using System;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace Rubberduck.UI.Settings
{
    public class IndenterSettingsViewModel : ViewModelBase, ISettingsViewModel
    {
        public IndenterSettingsViewModel(Configuration config)
        {
            _alignCommentsWithCode = config.UserSettings.IndenterSettings.AlignCommentsWithCode;
            _alignContinuations = config.UserSettings.IndenterSettings.AlignContinuations;
            _alignDimColumn = config.UserSettings.IndenterSettings.AlignDimColumn;
            _alignDims = config.UserSettings.IndenterSettings.AlignDims;
            _enableUndo = config.UserSettings.IndenterSettings.EnableUndo;
            _endOfLineCommentColumnSpaceAlignment = config.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment;
            _endOfLineCommentStyle = config.UserSettings.IndenterSettings.EndOfLineCommentStyle;
            _forceCompilerDirectivesInColumn1 = config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1;
            _forceDebugStatementsInColumn1 = config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1;
            _ignoreOperatorsInContinuations = config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations;
            _indentCase = config.UserSettings.IndenterSettings.IndentCase;
            _indentCompilerDirectives = config.UserSettings.IndenterSettings.IndentCompilerDirectives;
            _indentEntireProcedureBody = config.UserSettings.IndenterSettings.IndentEntireProcedureBody;
            _indentFirstCommentBlock = config.UserSettings.IndenterSettings.IndentFirstCommentBlock;
            _indentFirstDeclarationBlock = config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock;
            _indentSpaces = config.UserSettings.IndenterSettings.IndentSpaces;

            PropertyChanged += IndenterSettingsViewModel_PropertyChanged;
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

        private bool _enableUndo;
        public bool EnableUndo
        {
            get { return _enableUndo; }
            set
            {
                if (_enableUndo != value)
                {
                    _enableUndo = value;
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

        // ReSharper disable once InconsistentNaming
        private const string _previewSampleCode =
@"' Example Procedure
Sub ExampleProc()

    ' SMART INDENTER
    '@ 2016 by Rubberduck VBA.

    Dim iCount As Integer
    Static sName As String

    If YouWantMoreExamplesAndTools Then
        ' Visit http://rubberduckvba.com/

        Select Case X
        Case ""A""
            ' If you have any comments or suggestions, _
	         or find valid VBA code that isn't indented correctly,

	        #If VBA6 Then
	            MsgBox ""Contact contact@rubberduck-vba.com""
	        #End If

        Case ""Continued strings and parameters can be"" _
           & ""lined up for easier reading, optionally ignoring"" _
           , ""any operatirs (&+, etc) at the start of the line.""

           Debug.Print ""X<>1""
        End Select                                'Case X
    End If                                        'More Tools?

End Sub
";

        public string PreviewSampleCode 
        {
            get
            {
                var indenter = new Indenter(null, GetCurrentSettings);

                var lines = _previewSampleCode.Split(new[] {Environment.NewLine}, StringSplitOptions.None);
                indenter.Indent(lines, "TestModule", false);
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
                EnableUndo = EnableUndo,
                EndOfLineCommentColumnSpaceAlignment = EndOfLineCommentColumnSpaceAlignment,
                EndOfLineCommentStyle = EndOfLineCommentStyle,
                ForceCompilerDirectivesInColumn1 = ForceCompilerDirectivesInColumn1,
                ForceDebugStatementsInColumn1 = ForceDebugStatementsInColumn1,
                IgnoreOperatorsInContinuations = IgnoreOperatorsInContinuations,
                IndentCase = IndentCase,
                IndentCompilerDirectives = IndentCompilerDirectives,
                IndentEntireProcedureBody = IndentEntireProcedureBody,
                IndentFirstCommentBlock = IndentFirstCommentBlock,
                IndentFirstDeclarationBlock = IndentFirstDeclarationBlock,
                IndentSpaces = IndentSpaces
            };
        }

        #endregion

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.IndenterSettings.AlignCommentsWithCode = AlignCommentsWithCode;
            config.UserSettings.IndenterSettings.AlignContinuations = AlignContinuations;
            config.UserSettings.IndenterSettings.AlignDimColumn = AlignDimColumn;
            config.UserSettings.IndenterSettings.AlignDims = AlignDims;
            config.UserSettings.IndenterSettings.EnableUndo = EnableUndo;
            config.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment = EndOfLineCommentColumnSpaceAlignment;
            config.UserSettings.IndenterSettings.EndOfLineCommentStyle = EndOfLineCommentStyle;
            config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1 = ForceCompilerDirectivesInColumn1;
            config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1 = ForceDebugStatementsInColumn1;
            config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations = IgnoreOperatorsInContinuations;
            config.UserSettings.IndenterSettings.IndentCase = IndentCase;
            config.UserSettings.IndenterSettings.IndentCompilerDirectives = IndentCompilerDirectives;
            config.UserSettings.IndenterSettings.IndentEntireProcedureBody = IndentEntireProcedureBody;
            config.UserSettings.IndenterSettings.IndentFirstCommentBlock = IndentFirstCommentBlock;
            config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock = IndentFirstDeclarationBlock;
            config.UserSettings.IndenterSettings.IndentSpaces = IndentSpaces;
        }

        public void SetToDefaults(Configuration config)
        {
            AlignCommentsWithCode = config.UserSettings.IndenterSettings.AlignCommentsWithCode;
            AlignContinuations = config.UserSettings.IndenterSettings.AlignContinuations;
            AlignDimColumn = config.UserSettings.IndenterSettings.AlignDimColumn;
            AlignDims = config.UserSettings.IndenterSettings.AlignDims;
            EnableUndo = config.UserSettings.IndenterSettings.EnableUndo;
            EndOfLineCommentColumnSpaceAlignment = config.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment;
            EndOfLineCommentStyle = config.UserSettings.IndenterSettings.EndOfLineCommentStyle;
            ForceCompilerDirectivesInColumn1 = config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1;
            ForceDebugStatementsInColumn1 = config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1;
            IgnoreOperatorsInContinuations = config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations;
            IndentCase = config.UserSettings.IndenterSettings.IndentCase;
            IndentCompilerDirectives = config.UserSettings.IndenterSettings.IndentCompilerDirectives;
            IndentEntireProcedureBody = config.UserSettings.IndenterSettings.IndentEntireProcedureBody;
            IndentFirstCommentBlock = config.UserSettings.IndenterSettings.IndentFirstCommentBlock;
            IndentFirstDeclarationBlock = config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock;
            IndentSpaces = config.UserSettings.IndenterSettings.IndentSpaces;
        }
    }
}
