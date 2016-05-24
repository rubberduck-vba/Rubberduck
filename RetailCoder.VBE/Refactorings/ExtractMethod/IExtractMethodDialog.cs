using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodDialog : IDialogView
    {
        IEnumerable<ExtractedParameter> ReturnValues { get; set; }

        event EventHandler RefreshPreview;
        void OnRefreshPreview();
        string Preview { get; set; }

        string MethodName { get; set; }
        string OldMethodName { get; set; }
        Accessibility Accessibility { get; set; }
        IEnumerable<ExtractedParameter> Parameters { get; set; }
        IEnumerable<ExtractedParameter> Inputs { get; set; }
        IEnumerable<ExtractedParameter> Outputs { get; set; }
        IEnumerable<ExtractedParameter> Locals { get; set; }
    }
}
