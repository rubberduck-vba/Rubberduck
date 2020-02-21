using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter.RewriterInfo;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.Parsing.Rewriter
{
    public class ModuleRewriter : IExecutableModuleRewriter
    {
        private readonly QualifiedModuleName _module;
        private readonly ISourceCodeHandler _sourceCodeHandler;
        private readonly TokenStreamRewriter _rewriter;

        public ModuleRewriter(QualifiedModuleName module, ITokenStream tokenStream, ISourceCodeHandler sourceCodeHandler)
        {
            _module = module;
            _rewriter = new TokenStreamRewriter(tokenStream);
            _sourceCodeHandler = sourceCodeHandler;
        }

        public bool IsDirty => _rewriter.GetText() != _sourceCodeHandler.SourceCode(_module);

        public void Rewrite()
        {
            if (!IsDirty)
            {
                return;
            }

            var tentativeCode = _rewriter.GetText();

            while (tentativeCode.EndsWith(Environment.NewLine + Environment.NewLine))
            {
                tentativeCode = tentativeCode.Remove(tentativeCode.Length - Environment.NewLine.Length);
            };
            var newCode = tentativeCode;

            _sourceCodeHandler.SubstituteCode(_module, newCode);
        }

        public ITokenStream TokenStream => _rewriter.TokenStream;

        private static readonly IDictionary<Type, IRewriterInfoFinder> Finders =
            new Dictionary<Type, IRewriterInfoFinder>
            {
                {typeof(VBAParser.VariableSubStmtContext), new VariableRewriterInfoFinder()},
                {typeof(VBAParser.ConstSubStmtContext), new ConstantRewriterInfoFinder()},
                {typeof(VBAParser.ArgContext), new ParameterRewriterInfoFinder()},
                {typeof(VBAParser.ArgumentContext), new ArgumentRewriterInfoFinder()},
            };

        public void Remove(Declaration target)
        {
            Remove(target.Context);
        }

        public void Remove(ParserRuleContext target)
        {
            var info = Finders.TryGetValue(target.GetType(), out var finder)
                ? finder.GetRewriterInfo(target)
                : new DefaultRewriterInfoFinder().GetRewriterInfo(target);

            if (info.Equals(RewriterInfo.RewriterInfo.None)) { return; }
            _rewriter.Delete(info.StartTokenIndex, info.StopTokenIndex);
        }

        public void Remove(ITerminalNode target)
        {
            _rewriter.Delete(target.Symbol.TokenIndex);
        }

        public void Remove(IToken target)
        {
            _rewriter.Delete(target);
        }

        public void Remove(IParseTree target)
        {
            switch (target)
            {
                case ITerminalNode terminalNode:
                    Remove(terminalNode);
                    break;
                case ParserRuleContext context:
                    Remove(context);
                    break;
                default:
                    //It should be impossible to end up here.
                    throw new NotSupportedException();
            }
        }

        public void RemoveRange(int start, int stop)
        {
            _rewriter.Delete(start, stop);
        }

        public void Replace(Declaration target, string content)
        {
            _rewriter.Replace(target.Context.Start.TokenIndex, target.Context.Stop.TokenIndex, content);
        }

        public void Replace(ParserRuleContext target, string content)
        {
            _rewriter.Replace(target.Start.TokenIndex, target.Stop.TokenIndex, content);
        }

        public void Replace(IToken token, string content)
        {
            _rewriter.Replace(token, content);
        }

        public void Replace(ITerminalNode target, string content)
        {
            _rewriter.Replace(target.Symbol.TokenIndex, content);
        }

        public void Replace(IParseTree target, string content)
        {
            switch (target)
            {
                case ITerminalNode terminalNode:
                    Replace(terminalNode, content);
                    break;
                case ParserRuleContext context:
                    Replace(context, content);
                    break;
                default:
                    //It should be impossible to end up here.
                    throw new NotSupportedException();
            }
        }

        public void Replace(Interval tokenInterval, string content)
        {
            _rewriter.Replace(tokenInterval.a, tokenInterval.b, content);
        }

        public void InsertBefore(int tokenIndex, string content)
        {
            _rewriter.InsertBefore(tokenIndex, content);
        }

        public void InsertAfter(int tokenIndex, string content)
        {
            _rewriter.InsertAfter(tokenIndex, content);
        }

        public Selection? Selection { get; set; }
        public Selection? SelectionOffset { get; set; }

        public string GetText(int startTokenIndex, int stopTokenIndex)
        {
            return _rewriter.GetText(Interval.Of(startTokenIndex, stopTokenIndex));
        }

        public string GetText()
        {
            return _rewriter.GetText();
        }
    }
}