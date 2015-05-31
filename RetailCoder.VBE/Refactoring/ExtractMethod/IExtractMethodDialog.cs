using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactoring.ExtractMethod
{
    public interface IExtractMethodDialog : IDialogView
    {
        IEnumerable<ExtractedParameter> ReturnValues { get; set; }
        bool CanSetReturnValue { get; set; }

        event EventHandler RefreshPreview;
        void OnRefreshPreview();
        string Preview { get; set; }

        string MethodName { get; set; }
        Accessibility Accessibility { get; set; }
        bool SetReturnValue { get; set; }
        ExtractedParameter ReturnValue { get; set; }
        IEnumerable<ExtractedParameter> Parameters { get; set; }
        IEnumerable<ExtractedParameter> Inputs { get; set; }
        IEnumerable<ExtractedParameter> Outputs { get; set; }
        IEnumerable<ExtractedParameter> Locals { get; set; }
    }
}