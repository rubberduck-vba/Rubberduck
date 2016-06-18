using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.RegexAssistant
{
    public interface IRegularExpression
    {
        Quantifier Quantifier { get; }


        String Description { get; }
        Boolean TryMatch(ref String text);
        
    }
}
