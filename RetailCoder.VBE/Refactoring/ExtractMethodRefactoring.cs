using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Rubberduck.Reflection;

namespace Rubberduck.Refactoring
{
    [ComVisible(false)]
    public class ExtractMethodRefactoring : IRefactoring
    {
        public void Refactor(CodeModule module)
        {
            var codePane = module.CodePane;
            var selection = codePane.GetSelection();
            var scope = codePane.SelectedProcedure(selection);

            if (!IsValidSelection(module, selection))
            {
                return;
            }

            /* find all referenced identifiers in selection.
             * find all referenced identifiers in containing procedure.
             * if all identifiers referenced in selection are only used in selection, extract selection into a [Private Sub].
             * if idenfitiers referenced in selection are also used after selection, we need a [ByRef parameter] for each.
             * ...but if there's only one, we make a [Private Function] instead, and return the value.
             * 
             */
        }

        private bool IsValidSelection(CodeModule module, Selection selection)
        {
            if (selection.LineCount > 1)
            {
                vbext_ProcKind kindStart;
                var startProc = module.get_ProcOfLine(selection.StartLine, out kindStart);

                vbext_ProcKind kindEnd;
                var endProc = module.get_ProcOfLine(selection.EndLine, out kindEnd);

                if (startProc != endProc || kindStart != kindEnd)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
