using System.Collections.Generic;

namespace Rubberduck.RegexAssistant
{
    public interface IRegularExpression : IDescribable
    {
        IList<IRegularExpression> Subexpressions { get; }
    }
}