using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodExtraction : IExtractMethodExtraction
    {

        public void Apply(ICodeModule codeModule, IExtractMethodModel model, Selection selection)
        {
            var newMethodCall = model.Method.NewMethodCall();
            var positionToInsertNewMethod = model.PositionForNewMethod;
            var positionForMethodCall = model.PositionForMethodCall;
            var selectionToRemove = model.RowsToRemove;
            // The next 4 lines are dependent on the positions of the various parts,
            // so have to be applied in the correct order.
            var newMethod = ConstructLinesOfProc(codeModule, model);
            codeModule.InsertLines(positionToInsertNewMethod.StartLine, newMethod);
            RemoveSelection(codeModule, selectionToRemove);
            codeModule.InsertLines(selection.StartLine, newMethodCall);
        }

        public virtual void RemoveSelection(ICodeModule codeModule, IEnumerable<Selection> selection)
        {
            foreach (var item in selection.OrderBy(x => -x.StartLine))
            {
                var start = item.StartLine;
                var end = item.EndLine;
                var lineCount = end - start + 1;
                codeModule.DeleteLines(start,lineCount);
            }
        }

        public virtual string ConstructLinesOfProc(ICodeModule codeModule, IExtractMethodModel model)
        {

            var newLine = Environment.NewLine;
            var method = model.Method;
            var keyword = Tokens.Sub;
            var asTypeClause = string.Empty;
            var selection = model.RowsToRemove;

            var access = method.Accessibility.ToString();
            var extractedParams = method.Parameters.Select(p => ExtractedParameter.PassedBy.ByRef + " " + p.Name + " " + Tokens.As + " " + p.TypeName);
            var parameters = "(" + string.Join(", ", extractedParams) + ")";
            //method signature
            var result = access + ' ' + keyword + ' ' + method.MethodName + parameters + ' ' + asTypeClause + newLine;
            // method body
            string textToMove = "";
            foreach (var item in selection)
            {
                textToMove += codeModule.GetLines(item.StartLine, item.EndLine - item.StartLine + 1);
                textToMove += Environment.NewLine;
            }
            // method end;
            result += textToMove;
            result += Tokens.End + " " + Tokens.Sub;
            return result;
        }
    }
}
