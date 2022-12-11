using NLog;
using Rubberduck.Interaction;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.Refactorings.ExtractMethod
{
    internal class ExtractMethodPresenter : RefactoringPresenterBase<ExtractMethodModel>, IExtractMethodPresenter
    {
        private static readonly DialogData DialogData = DialogData.Create(RefactoringsUI.ExtractMethod_Caption, 600, 600); //TODO - get appropriate size

        private readonly IMessageBox _messageBox;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ExtractMethodPresenter(ExtractMethodModel model,
            IRefactoringDialogFactory dialogFactory, IMessageBox messageBox) : 
            base(DialogData, model, dialogFactory)
        {
            _messageBox = messageBox;
        }

        override public ExtractMethodModel Show()
        {
            //TODO - test not cancelled or other invalid output?
            return base.Show();
        }

        //ExtractMethodModel IExtractMethodPresenter.Show(IExtractMethodModel methodModel, IExtractMethodProc extractMethodProc)
        //{
        //    throw new System.NotImplementedException();
        //}
        /*
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
