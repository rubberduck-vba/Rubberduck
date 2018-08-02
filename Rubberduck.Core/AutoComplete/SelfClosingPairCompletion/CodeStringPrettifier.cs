using Rubberduck.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public class CodeStringPrettifier : ICodeStringPrettifier
    {
        public CodeString Run(CodeString code, ICodeModule module)
        {
            using (var pane = module.CodePane)
            {
                using (var window = pane.Window)
                {
                    window.ScreenUpdating = false;
                    module.DeleteLines(code.SnippetPosition);
                    module.InsertLines(code.SnippetPosition.StartLine, code.Code);
                    pane.Selection = new Selection(code.SnippetPosition.StartLine, 1).Offset(code.CaretPosition);
                    window.ScreenUpdating = true;

                    var lines = module.GetLines(code.SnippetPosition);
                    if (lines.Equals(code.Code, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return code;
                    }
                }
            }

            return default;
        }
    }
}
