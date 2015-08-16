using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodPresenter
    {
        private readonly IExtractMethodDialog _view;
        private readonly ExtractMethodModel _model;

        public ExtractMethodPresenter(IExtractMethodDialog view, ExtractMethodModel model)
        {
            _view = view;
            _model = model;
        }

        public ExtractMethodModel Show()
        {
            PrepareView();

            if (_view.ShowDialog() != DialogResult.OK)
            {
                return null;
            }

            return _model;
        }

        private void PrepareView()
        {
            _view.MethodName = _model.SourceMember.IdentifierName + "_1";
            _view.Inputs = _model.Inputs;
            _view.Outputs = _model.Outputs;
            _view.Locals =
                _model.Locals.Select(
                    variable =>
                        new ExtractedParameter(variable.AsTypeName, ExtractedParameter.PassedBy.ByVal, variable.IdentifierName))
                    .ToList();

            var returnValues = new[] {new ExtractedParameter(string.Empty, ExtractedParameter.PassedBy.ByVal)}
                .Union(_view.Outputs)
                .Union(_view.Inputs)
                .ToList();

            _view.ReturnValues = returnValues;
            _view.ReturnValue = (_model.Outputs.Any() && !_model.Outputs.Skip(1).Any())
                ? _model.Outputs.Single()
                : returnValues.First();

            _view.RefreshPreview += _view_RefreshPreview;
            _view.OnRefreshPreview();
        }

        private static readonly IEnumerable<string> ValueTypes = new[]
        {
            Tokens.Boolean,
            Tokens.Byte,
            Tokens.Currency,
            Tokens.Date,
            Tokens.Decimal,
            Tokens.Double,
            Tokens.Integer,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.Single,
            Tokens.String
        };

        private void _view_RefreshPreview(object sender, EventArgs e)
        {
            var hasReturnValue = _model.Method.ReturnValue != null && _model.Method.ReturnValue.Name != ExtractedParameter.None;
            _view.CanSetReturnValue =
                hasReturnValue && !ValueTypes.Contains(_model.Method.ReturnValue.TypeName);

            GeneratePreview();
        }

        private void GeneratePreview()
        {
            _model.Method.MethodName = _view.MethodName;
            _model.Method.Accessibility = _view.Accessibility;
            _model.Method.Parameters = _view.Parameters;
            _model.Method.ReturnValue = _view.ReturnValue;
            _model.Method.SetReturnValue = _view.SetReturnValue;

            _view.Preview = ExtractMethodRefactoring.GetExtractedMethod(_model);
        }
    }
}
