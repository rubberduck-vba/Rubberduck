using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter.RewriterInfo;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Rewriter
{
    public class ModuleRewriter : IModuleRewriter
    {
        protected ICodeModule Module { get; }
        protected TokenStreamRewriter Rewriter { get; }

        public ModuleRewriter(ICodeModule module, TokenStreamRewriter rewriter)
        {
            Module = module;
            Rewriter = rewriter;
        }

        public virtual bool IsDirty { get; private set; }
        public ITokenStream TokenStream => Rewriter.TokenStream;

        public virtual void Rewrite()
        {
            if (!IsDirty) { return; }

            Module.Clear();
            var content = Rewriter.GetText();
            Module.InsertLines(1, content);
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
            Rewriter.Delete(info.StartTokenIndex, info.StopTokenIndex);
            IsDirty = true;
        }

        public void Remove(ITerminalNode target)
        {
            Rewriter.Delete(target.Symbol.TokenIndex);
            IsDirty = true;
        }

        public void Remove(IToken target)
        {
            Rewriter.Delete(target);
            IsDirty = true;
        }

        public void Replace(Declaration target, string content)
        {
            Rewriter.Replace(target.Context.Start.TokenIndex, target.Context.Stop.TokenIndex, content);
            IsDirty = true;
        }

        public void Replace(ParserRuleContext target, string content)
        {
            Rewriter.Replace(target.Start.TokenIndex, target.Stop.TokenIndex, content);
            IsDirty = true;
        }

        public void Replace(IToken token, string content)
        {
            Rewriter.Replace(token, content);
            IsDirty = true;
        }

        public void Replace(ITerminalNode target, string content)
        {
            Rewriter.Replace(target.Symbol.TokenIndex, content);
            IsDirty = true;
        }

        public void Replace(Interval tokenInterval, string content)
        {
            Rewriter.Replace(tokenInterval.a, tokenInterval.b, content);
            IsDirty = true;
        }

        public void InsertBefore(int tokenIndex, string content)
        {
            Rewriter.InsertBefore(tokenIndex, content);
            IsDirty = true;
        }

        public void InsertAfter(int tokenIndex, string content)
        {
            Rewriter.InsertAfter(tokenIndex, content);
            IsDirty = true;
        }

        public string GetText(int startTokenIndex, int stopTokenIndex)
        {
            return Rewriter.GetText(Interval.Of(startTokenIndex, stopTokenIndex));
        }

        public string GetText()
        {
            return Rewriter.GetText();
        }
    }
}