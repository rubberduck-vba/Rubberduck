using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    public interface IExtractMethodDialog : IDialogView
    {
        ExtractedParameter ReturnValue { get; set; }
        IEnumerable<ExtractedParameter> ReturnValues { get; set; }

        Accessibility Accessibility { get; set; }
        string MethodName { get; set; }
        bool SetReturnValue { get; set; }
        bool CanSetReturnValue { get; set; }

        event EventHandler RefreshPreview;
        void OnRefreshPreview();
        string Preview { get; set; }

        IEnumerable<ExtractedParameter> Parameters { get; set; }

        IEnumerable<ExtractedParameter> Inputs { get; set; }
        IEnumerable<ExtractedParameter> Outputs { get; set; }
        IEnumerable<ExtractedParameter> Locals { get; set; }
    }
}