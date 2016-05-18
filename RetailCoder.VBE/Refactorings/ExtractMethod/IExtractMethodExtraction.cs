using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodExtraction
    {
        void apply(IExtractMethodModel model, Selection selection);
    }
}
