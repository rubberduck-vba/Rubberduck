using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public static class IEnumerableExt
    {
        /// <summary>
        /// Yields an Enumeration of selector Type, 
        /// by checking for gaps between elements 
        /// using the supplied increment function to work out the next value
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="U"></typeparam>
        /// <param name="inputs"></param>
        /// <param name="getIncr"></param>
        /// <param name="selector"></param>
        /// <param name="comparisonFunc"></param>
        /// <returns></returns>
        public static IEnumerable<U> GroupByMissing<T, U>(this IEnumerable<T> inputs, Func<T, T> getIncr, Func<T, T, U> selector, Func<T, T, int> comparisonFunc)
        {

            var initialized = false;
            T first = default(T);
            T last = default(T);
            T next = default(T);
            Tuple<T, T> tuple = null;

            foreach (var input in inputs)
            {
                if (!initialized)
                {
                    first = input;
                    last = input;
                    initialized = true;
                    continue;
                }
                if (comparisonFunc(last, input) < 0)
                {
                    throw new ArgumentException(string.Format("Values are not monotonically increasing. {0} should be less than {1}", last, input));
                }
                var inc = getIncr(last);
                if (!input.Equals(inc))
                {
                    yield return selector(first, last);
                    first = input;
                }
                last = input;
            }
            if (initialized)
            {
                yield return selector(first, last);
            }
        }
    }

    public class ExtractMethodModel : IExtractMethodModel
    {
        private List<Declaration> _extractDeclarations;
        private IExtractMethodParameterClassification _paramClassify;
        private IExtractedMethod _extractedMethod;

        public ExtractMethodModel(IExtractedMethod extractedMethod, IExtractMethodParameterClassification paramClassify)
        {
            _extractedMethod = extractedMethod;
            _paramClassify = paramClassify;
        }

        public void extract(IEnumerable<Declaration> declarations, QualifiedSelection selection, string selectedCode)
        {
            var items = declarations.ToList();
            _selection = selection;
            _selectedCode = selectedCode;
            _rowsToRemove = new List<Selection>();

            var sourceMember = items.FindSelectedDeclaration(
                selection,
                DeclarationExtensions.ProcedureTypes,
                d => ((ParserRuleContext)d.Context.Parent).GetSelection());

            if (sourceMember == null)
            {
                throw new InvalidOperationException("Invalid selection.");
            }

            var inScopeDeclarations = items.Where(item => item.ParentScope == sourceMember.Scope).ToList();
            var selectionStartLine = selection.Selection.StartLine;
            var selectionEndLine = selection.Selection.EndLine;
            var methodInsertLine = sourceMember.Context.Stop.Line + 1;

            _positionForNewMethod = new Selection(methodInsertLine, 1, methodInsertLine, 1);

            foreach (var item in inScopeDeclarations)
            {
                _paramClassify.classifyDeclarations(selection, item);
            }
            _declarationsToMove = _paramClassify.DeclarationsToMove.ToList();

            _rowsToRemove = splitSelection(selection.Selection, _declarationsToMove).ToList();

            var methodCallPositionStartLine = selectionStartLine - _declarationsToMove.Count(d => d.Selection.StartLine < selectionStartLine);
            _positionForMethodCall = new Selection(methodCallPositionStartLine, 1, methodCallPositionStartLine, 1);
            _extractedMethod.ReturnValue = null;
            _extractedMethod.Accessibility = Accessibility.Private;
            _extractedMethod.SetReturnValue = false;
            _extractedMethod.Parameters = _paramClassify.ExtractedParameters.ToList();

        }

        public IEnumerable<Selection> splitSelection(Selection selection, IEnumerable<Declaration> declarations)
        {
            var tupleList = new List<Tuple<int, int>>();
            var declarationRows = declarations
                .Where(decl =>
                    selection.StartLine <= decl.Selection.StartLine &&
                    decl.Selection.StartLine <= selection.EndLine)
                .Select(decl => decl.Selection.StartLine)
                .OrderBy(x => x)
                .ToList();

            var gappedSelectionRows = Enumerable.Range(selection.StartLine, selection.EndLine - selection.StartLine + 1).Except(declarationRows).ToList();
            var returnList = gappedSelectionRows.GroupByMissing(x => (x + 1), (x, y) => new Selection(x, 1, y, 1), (x, y) => y - x);
            return returnList;
        }

        private Declaration _sourceMember;
        public Declaration SourceMember { get { return _sourceMember; } }

        private QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        private string _selectedCode;
        public string SelectedCode { get { return _selectedCode; } }

        private List<Declaration> _locals;
        public IEnumerable<Declaration> Locals { get { return _locals; } }

        private IEnumerable<ExtractedParameter> _input;
        public IEnumerable<ExtractedParameter> Inputs { get { return _input; } }
        private IEnumerable<ExtractedParameter> _output;
        public IEnumerable<ExtractedParameter> Outputs { get { return _output; } }

        private List<Declaration> _declarationsToMove;
        public IEnumerable<Declaration> DeclarationsToMove { get { return _declarationsToMove; } }

        public IExtractedMethod Method { get { return _extractedMethod; } }

        private Selection _positionForMethodCall;
        public Selection PositionForMethodCall { get { return _positionForMethodCall; } }

        public string NewMethodCall { get { return _extractedMethod.NewMethodCall(); } }

        private Selection _positionForNewMethod;
        public Selection PositionForNewMethod { get { return _positionForNewMethod; } }
        IList<Selection> _rowsToRemove;
        public IEnumerable<Selection> RowsToRemove
        {
            // we need to split selectionToRemove around any declarations that
            // are within the selection.
            get { return _declarationsToMove.Select(decl => decl.Selection).Union(_rowsToRemove)
                .Select( x => new Selection(x.StartLine,1,x.EndLine,1)) ; }
        }

        public IEnumerable<Declaration> DeclarationsToExtract
        {
            get { return _extractDeclarations; }
        }
    }
}