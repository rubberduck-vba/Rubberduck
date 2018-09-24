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

            return new CodeString(prettifiedCode, new Selection(0, index + 1));
        }
    }
}
