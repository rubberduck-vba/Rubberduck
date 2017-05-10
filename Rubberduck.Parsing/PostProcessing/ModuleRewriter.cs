using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PostProcessing.RewriterInfo;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.PostProcessing
{
    public class ModuleRewriter : IModuleRewriter
    {
        private readonly ICodeModule _module;
        private readonly TokenStreamRewriter _rewriter;

        public ITokenStream TokenStream => _rewriter.TokenStream;

        public ModuleRewriter(ICodeModule module, TokenStreamRewriter rewriter)
        {
            _module = module;
            _rewriter = rewriter;
        }

        public void Rewrite()
        {
            _module.Clear();
            var content = _rewriter.GetText();
            _module.InsertLines(1, content);
        }

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
            IRewriterInfoFinder finder;
            var info = Finders.TryGetValue(target.GetType(), out finder)
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