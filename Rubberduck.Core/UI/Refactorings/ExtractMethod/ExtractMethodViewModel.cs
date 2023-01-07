using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractMethod;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    public class ExtractMethodViewModel : RefactoringViewModelBase<ExtractMethodModel>
    {
        private readonly IMessageBox _messageBox;

        public RubberduckParserState State { get; }

        public ExtractMethodViewModel(RubberduckParserState state, ExtractMethodModel model, IMessageBox messageBox)
            : base(model)
        {
            State = state;
            _messageBox = messageBox;
            //_model = model;
        }

        private bool _wired;
        public ObservableCollection<ExtractMethodParameter> Parameters
        {
            get
            {
                if (!_wired)
                {
                    WireParameterEvents();
                }
                return Model.Parameters;
            }
            set
            {
                Model.Parameters = value;
                WireParameterEvents();
                OnPropertyChanged(nameof(PreviewCode));
                OnPropertyChanged(nameof(ReturnParameters));
                OnPropertyChanged(nameof(ReturnParameter));
            }
        }

        private void WireParameterEvents()
        {
            foreach (var parameter in Model.Parameters)
            {
                parameter.PropertyChanged += Parameter_PropertyChanged;
            }
            _wired = true;
        }

        private void Parameter_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            OnPropertyChanged(nameof(PreviewCode));
        }

        public IEnumerable<string> ComponentNames => Model.ComponentNames;

        public string NewMethodName
        {
            get => Model.NewMethodName;
            set
            {
                Model.NewMethodName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidMethodName));
                OnPropertyChanged(nameof(PreviewCode));
            }
        }

        public IEnumerable<ExtractMethodParameter> ReturnParameters =>
            new[]
            {
                ExtractMethodParameter.None
            }.Union(Parameters.Where(p => p.CanReturn));

        public ExtractMethodParameter ReturnParameter
        {
            get => Model.ReturnParameter ?? ExtractMethodParameter.None;
            set
            {
                Model.ReturnParameter = value ?? ExtractMethodParameter.None;
                OnPropertyChanged(nameof(PreviewCode));
            }
        }

        public string SourceMethodName => Model.SourceMethodName;
        public string PreviewCaption => string.Format(RefactoringsUI.ExtractMethod_CodePreviewCaption, SourceMethodName);
        public string PreviewCode => Model.NewMethodCode; //Model.ReplacementCode + Environment.NewLine + Model.PreviewCode; //any way to get a divider for showing both sections of code in one textbox?
        public IEnumerable<ExtractMethodParameter> Inputs;
        public IEnumerable<ExtractMethodParameter> Outputs;
        public IEnumerable<ExtractMethodParameter> Locals;
        public IEnumerable<ExtractMethodParameter> ReturnValues;
        public Accessibility Accessibility;

        public bool IsValidMethodName
        {
            get
            {
                var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);
                return !string.IsNullOrWhiteSpace(NewMethodName)
                       && char.IsLetter(NewMethodName.FirstOrDefault())
                       && !NewMethodName.Any(c => !char.IsLetterOrDigit(c) && c != '_')
                       && !ComponentNames.Contains(NewMethodName, StringComparer.InvariantCultureIgnoreCase)
                       && !tokenValues.Contains(NewMethodName, StringComparer.InvariantCultureIgnoreCase);
            }
        }

        public bool DisplayCompilationConstantWarning => !Model.ModuleContainsCompilationDirectives;

        protected override void DialogOk()
        {
            base.DialogOk();
        }
    }
}
