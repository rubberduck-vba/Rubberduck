using Rubberduck.Common;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public class CodeStringPrettifier : ICodeStringPrettifier
    {
        private readonly ICodeModule _module;

        public CodeStringPrettifier(ICodeModule module)
        {
            _module = module;
        }

        public bool IsSpacingUnchanged(CodeString code, CodeString original)
        {
            using (var pane = _module.CodePane)
            {
                using (var window = pane.Window)
                {
                    //window.ScreenUpdating = false;
                    _module.DeleteLines(code.SnippetPosition);
                    _module.InsertLines(code.SnippetPosition.StartLine, code.Code);
                    //window.ScreenUpdating = true;

                    pane.Selection = code.SnippetPosition.Offset(code.CaretPosition);

                    var lines = _module.GetLines(code.SnippetPosition);
                    if (lines.Equals(code.Code, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return true;
                    }
                }

                _module.ReplaceLine(code.SnippetPosition.StartLine, original.Code);
            }
            return false;
        }
    }
}
