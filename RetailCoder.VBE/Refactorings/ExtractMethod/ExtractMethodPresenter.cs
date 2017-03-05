using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodPresenter
    {
        bool Show(IExtractMethodModel model, IExtractMethodProc createProc);
    }

    public class ExtractMethodPresenter : IExtractMethodPresenter
    {
        private readonly IExtractMethodDialog _view;
        private readonly IIndenter _indenter;

        public ExtractMethodPresenter(IExtractMethodDialog view, IIndenter indenter)
        {
            _view = view;
            _indenter = indenter;
        }

        public bool Show(IExtractMethodModel methodModel, IExtractMethodProc extractMethodProc)
        {
            PrepareView(methodModel,extractMethodProc);

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return false;
            }

            return true;
        }

        private void PrepareView(IExtractMethodModel extractedMethodModel, IExtractMethodProc extractMethodProc)
        {
            _view.OldMethodName = extractedMethodModel.SourceMember.IdentifierName;
            _view.MethodName = extractedMethodModel.SourceMember.IdentifierName + "_1";
            _view.Inputs = extractedMethodModel.Inputs;
            _view.Outputs = extractedMethodModel.Outputs;
            _view.Locals =
                extractedMethodModel.Locals.Select(
                    variable =>
                        new ExtractedParameter(variable.AsTypeName, ExtractedParameter.PassedBy.ByVal, variable.IdentifierName))
                    .ToList();

            var returnValues = new[] {new ExtractedParameter(string.Empty, ExtractedParameter.PassedBy.ByVal)}
                .Union(_view.Outputs)
                .Union(_view.Inputs)
                .ToList();

            _view.ReturnValues = returnValues;

            _view.RefreshPreview += (object sender, EventArgs e) => { GeneratePreview(extractedMethodModel, extractMethodProc); };

            _view.OnRefreshPreview();
        }

        private void GeneratePreview(IExtractMethodModel extractMethodModel,IExtractMethodProc extractMethodProc )
        {
            extractMethodModel.Method.MethodName = _view.MethodName;
            extractMethodModel.Method.Accessibility = _view.Accessibility;
            extractMethodModel.Method.Parameters = _view.Parameters;
            /*
             * extractMethodModel.Method.ReturnValue = _view.ReturnValue;
             * extractMethodModel.Method.SetReturnValue = _view.SetReturnValue;
             */
            var extractedMethod = extractMethodProc.createProc(extractMethodModel);
            var code = extractedMethod.Split(new[]{Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);
            code = _indenter.Indent(code).ToArray();
            _view.Preview = string.Join(Environment.NewLine, code);
        }
    }
}
