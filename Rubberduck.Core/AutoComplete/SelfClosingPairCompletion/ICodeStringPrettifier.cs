using Rubberduck.Common;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public interface ICodeStringPrettifier
    {
        CodeString Run(CodeString code, ICodeModule module);
    }
}
