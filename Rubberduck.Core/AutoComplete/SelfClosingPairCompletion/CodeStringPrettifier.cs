using Rubberduck.Common;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.VBEditor;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public class CodeStringPrettifier : ICodeStringPrettifier
    {
        private readonly ICodeModule _module;

        public CodeStringPrettifier(ICodeModule module)
        {
            _module = module;
        }

        public CodeString Prettify(CodeString original)
        {
            var originalCode = original.Code;
            var originalPosition = original.CaretPosition.StartColumn;
            var originalNonSpacePosition = 0;
            for (var i = 0; i < originalPosition; i++)
            {
                if (originalCode[i] != ' ')
                {
                    originalNonSpacePosition++;
                }
            }

            _module.DeleteLines(original.CaretPosition.StartLine);
            _module.InsertLines(original.CaretPosition.StartLine, originalCode);
            var prettifiedCode = _module.GetLines(original.CaretPosition);

            var prettifiedNonSpacePosition = 0;
            var index = 0;
            for (var i = 0; i < prettifiedCode.Length; i++)
            {
                if (prettifiedCode[i] != ' ')
                {
                    prettifiedNonSpacePosition++;
                    if (prettifiedNonSpacePosition == originalNonSpacePosition)
                    {
                        index = i;
                        break;
                    }
                }
            }

            return new CodeString(prettifiedCode, new Selection(0, index + 1));
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
