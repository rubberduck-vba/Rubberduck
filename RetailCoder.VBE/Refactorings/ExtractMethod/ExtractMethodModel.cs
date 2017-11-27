using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;

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

    public class ExtractMethodModel
    {
        public IEnumerable<ParserRuleContext> SelectedContexts { get; }
        public RubberduckParserState State { get; }
        public IIndenter Indenter { get; }
        public QualifiedSelection Selection { get; }

        public string OriginalMethodName { get; private set; }

        public ExtractMethodModel(RubberduckParserState state, QualifiedSelection selection,
            IEnumerable<ParserRuleContext> selectedContexts, IIndenter indenter)
        {
            State = state;
            Indenter = indenter;
            Selection = selection;
            SelectedContexts = selectedContexts;
            Setup();
        }

        private void Setup()
        {
            var topContext = SelectedContexts.First();
            ParserRuleContext stmtContext = null;
            var currentContext = (RuleContext)topContext;
            do {
                switch (currentContext)
                {
                    case VBAParser.FunctionStmtContext stmt:
                        stmtContext = stmt;
                        OriginalMethodName = stmt.functionName().GetText();
                        break;
                    case VBAParser.SubStmtContext stmt:
                        stmtContext = stmt;
                        OriginalMethodName = stmt.subroutineName().GetText();
                        break;
                    case VBAParser.PropertyGetStmtContext stmt:
                        stmtContext = stmt;
                        OriginalMethodName = stmt.functionName().GetText();
                        break;
                    case VBAParser.PropertyLetStmtContext stmt:
                        stmtContext = stmt;
                        OriginalMethodName = stmt.subroutineName().GetText();
                        break;
                    case VBAParser.PropertySetStmtContext stmt:
                        stmtContext = stmt;
                        OriginalMethodName = stmt.subroutineName().GetText();
                        break;
                }
                currentContext = currentContext.Parent;
            }
            while (currentContext != null && stmtContext == null) ;

            
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

            var gappedSelectionRows = Enumerable.Range(selection.StartLine, selection.EndLine - selection.StartLine + 1)
                .Except(declarationRows).ToList();
            var returnList =
                gappedSelectionRows.GroupByMissing(x => (x + 1), (x, y) => new Selection(x, 1, y, 1), (x, y) => y - x);
            return returnList;
        }

        public Declaration SourceMember
        {
            get => null;
        } // TODO: Remove from WPF

        private string _selectedCode;

        public string SelectedCode
        {
            get { return _selectedCode; }
        }

        private List<Declaration> _locals;

        public IEnumerable<Declaration> Locals
        {
            get { return _locals; }
        }

        private IEnumerable<ExtractedParameter> _input;

        public IEnumerable<ExtractedParameter> Inputs
        {
            get { return _input; }
        }

        private IEnumerable<ExtractedParameter> _output;

        public IEnumerable<ExtractedParameter> Outputs
        {
            get { return _output; }
        }

        private List<Declaration> _declarationsToMove;

        public IEnumerable<Declaration> DeclarationsToMove
        {
            get { return _declarationsToMove; }
        }

        //public IExtractedMethod Method { get { return _extractedMethod; } }

        private Selection _positionForMethodCall;

        public Selection PositionForMethodCall
        {
            get { return _positionForMethodCall; }
        }

        //public string NewMethodCall { get { return _extractedMethod.NewMethodCall(); } }

        private Selection _positionForNewMethod;

        public Selection PositionForNewMethod
        {
            get { return _positionForNewMethod; }
        }

        IList<Selection> _rowsToRemove;

        public IEnumerable<Selection> RowsToRemove
        {
            // we need to split selectionToRemove around any declarations that
            // are within the selection.
            get
            {
                return _declarationsToMove.Select(decl => decl.Selection).Union(_rowsToRemove)
                    .Select(x => new Selection(x.StartLine, 1, x.EndLine, 1));
            }
        }

        public IEnumerable<Declaration> DeclarationsToExtract
        {
            // TODO: Remove from WPF
            get => null;
        }
    }
}