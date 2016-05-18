using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodeModule;

namespace Rubberduck.Refactorings.ExtractMethod
{


    public class ExtractMethodExtraction : IExtractMethodExtraction
    {
        private readonly ICodeModuleWrapper _codeModule;
        private readonly IExtractMethodProc _createProc;

        public ExtractMethodExtraction(ICodeModuleWrapper codeModule, IExtractMethodProc createProc)
        {
            _codeModule = codeModule;
            _createProc = createProc;

        }
        public void apply(IExtractMethodModel model, Selection selection)
        {
            var newMethod = model.NewExtractedMethod(_createProc);
            var newMethodCall = model.Method.NewMethodCall();
            var positionToInsertNewMethod = model.PositionForNewMethod;
            var positionForMethodCall = model.PositionForMethodCall;
            var selectionToRemove = model.SelectionToRemove;

            // The next 4 lines are dependent on the positions of the various parts,
            // so have to be applied in the correct order.
            var textToMove = constructLinesOfProc(selectionToRemove);
            _codeModule.InsertLines(positionToInsertNewMethod.StartLine, newMethod);
            removeSelection(selectionToRemove);
            _codeModule.InsertLines(selection.StartLine, newMethodCall);
        }

        public void removeSelection(IEnumerable<Selection> selection)
        {
            foreach (var item in selection)
            {
                var start = item.StartLine;
                var end = item.EndLine;
                var lineCount = end - start + 1;
                var lineToDelete = new Selection(start, 1, start, 1);

                for (int i = 0; i < lineCount; i++)
                {
                    _codeModule.DeleteLines(lineToDelete);
                }

            }
        }
        public string constructLinesOfProc(IEnumerable<Selection> selection)
        {
            string textToMove = "";
            foreach (var item in selection)
            {
                textToMove += _codeModule.get_Lines(item.StartLine, item.EndLine - item.StartLine + 1);
                textToMove += Environment.NewLine;
            }
            return textToMove;
        }

    }
}
