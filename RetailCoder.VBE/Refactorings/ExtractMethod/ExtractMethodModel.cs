using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
        public ICodeModule CodeModule { get; }
        public QualifiedSelection Selection { get; }

        public string SourceMethodName { get; private set; }
        public string NewMethodName { get; set; }

        public ExtractMethodModel(RubberduckParserState state, QualifiedSelection selection,
            IEnumerable<ParserRuleContext> selectedContexts, IIndenter indenter, ICodeModule codeModule)
        {
            State = state;
            Indenter = indenter;
            CodeModule = codeModule;
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
                        SourceMethodName = stmt.functionName().GetText();
                        break;
                    case VBAParser.SubStmtContext stmt:
                        stmtContext = stmt;
                        SourceMethodName = stmt.subroutineName().GetText();
                        break;
                    case VBAParser.PropertyGetStmtContext stmt:
                        stmtContext = stmt;
                        SourceMethodName = stmt.functionName().GetText();
                        break;
                    case VBAParser.PropertyLetStmtContext stmt:
                        stmtContext = stmt;
                        SourceMethodName = stmt.subroutineName().GetText();
                        break;
                    case VBAParser.PropertySetStmtContext stmt:
                        stmtContext = stmt;
                        SourceMethodName = stmt.subroutineName().GetText();
                        break;
                }
                currentContext = currentContext.Parent;
            }
            while (currentContext != null && stmtContext == null) ;

            if (string.IsNullOrWhiteSpace(NewMethodName))
            {
                NewMethodName = "NewMethod";
            }

            SelectedCode = String.Join(Environment.NewLine, SelectedContexts.Select(c => c.GetText()));
        }
        
        public string SelectedCode { get; private set; }

        public string PreviewCode
        {
            get
            {
                //var rewriter = State.GetRewriter(CodeModule.GetQualifiedSelection().Value.QualifiedName);

                var strings = new List<string>();
                strings.Add($@"Public Sub {NewMethodName ?? "NewMethod"}");
                strings.AddRange(SelectedCode.Split(new[] {Environment.NewLine}, StringSplitOptions.None));
                strings.Add("End Sub");
                return string.Join(Environment.NewLine, Indenter.Indent(strings));
            }
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
    }
}