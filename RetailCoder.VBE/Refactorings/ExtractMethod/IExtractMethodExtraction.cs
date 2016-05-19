using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodeModule;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public interface IExtractMethodExtraction
    {
        void apply(ICodeModuleWrapper codeModule, IExtractMethodModel model, Selection selection);
    }
}
