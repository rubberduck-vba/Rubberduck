using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodProc
    {
        string createProc(IExtractMethodModel model);
    }
}
