using System;
using Rubberduck.Common;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
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

            _module.DeleteLines(original.SnippetPosition.StartLine);
            _module.InsertLines(original.SnippetPosition.StartLine, originalCode);
            var prettifiedCode = _module.GetLines(original.SnippetPosition);

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

            if (string.IsNullOrEmpty(original.Code) || original.Code[Math.Max(0, original.Code.Length - 1)] == ' ')
            {
                prettifiedCode += ' ';
            }
            var selection = new Selection(original.SnippetPosition.StartLine - 1, index).ToOneBased();
            using (var pane = _module.CodePane)
            {
                pane.Selection = selection;
            }
            return new CodeString(prettifiedCode, new Selection(0, selection.StartColumn - 1), selection);
        }
    }
}
