using Rubberduck.UI.Refactorings;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodPresenter : RefactoringPresenter<ExtractMethodModel, ExtractMethodDialog, ExtractMethodViewModel>
    {
        public ExtractMethodPresenter(ExtractMethodModel model, ExtractMethodDialog view, ExtractMethodViewModel viewModel) : base(model, view, viewModel) { }
        
        /*
        public ExtractMethodModel Show()
        {
            PrepareView(_model);

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            return _model;
        }

        private void PrepareView(ExtractMethodModel extractedMethodModel)
        {
            _view.ViewModel.OldMethodName = extractedMethodModel.SourceMember.IdentifierName;
            _view.ViewModel.MethodName = extractedMethodModel.SourceMember.IdentifierName;
            _view.ViewModel.Inputs = extractedMethodModel.Inputs;
            _view.ViewModel.Outputs = extractedMethodModel.Outputs;
            _view.ViewModel.Locals =
                extractedMethodModel.Locals.Select(
                    variable =>
                        new ExtractedParameter(variable.AsTypeName, PassedBy.ByVal, variable.IdentifierName))
                    .ToList();

            var returnValues = new[] {new ExtractedParameter(string.Empty, PassedBy.ByVal)}
                .Union(_view.ViewModel.Outputs)
                .Union(_view.ViewModel.Inputs)
                .ToList();

            _view.ViewModel.ReturnValues = returnValues;

            //_view.RefreshPreview += (object sender, EventArgs e) => { GeneratePreview(extractedMethodModel, extractMethodProc); };

            //_view.OnRefreshPreview();
        }

        private void GeneratePreview(IExtractMethodModel extractMethodModel,IExtractMethodProc extractMethodProc )
        {
            extractMethodModel.Method.MethodName = _view.ViewModel.MethodName;
            extractMethodModel.Method.Accessibility = _view.ViewModel.Accessibility;
            extractMethodModel.Method.Parameters = _view.ViewModel.Parameters;
             //
             // extractMethodModel.Method.ReturnValue = _view.ReturnValue;
             // extractMethodModel.Method.SetReturnValue = _view.SetReturnValue;
             //
            var extractedMethod = extractMethodProc.createProc(extractMethodModel);
            var code = extractedMethod.Split(new[]{Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);
            code = _indenter.Indent(code).ToArray();
            _view.ViewModel.Preview = string.Join(Environment.NewLine, code);
        }
        */
    }
}
