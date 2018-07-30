using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter.RewriterInfo;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Rewriter
{
    public class ModuleRewriter : IModuleRewriter
    {
        protected QualifiedModuleName Module { get; }
        protected IProjectsProvider ProjectsProvider { get; }
        protected TokenStreamRewriter Rewriter { get; }

        public ModuleRewriter(QualifiedModuleName module, ITokenStream tokenStream, IProjectsProvider projectsProvider)
        {
            Module = module;
            Rewriter = new TokenStreamRewriter(tokenStream);
            ProjectsProvider = projectsProvider;
        }

        /// <summary>
        /// Returns the code module of the module identified by Module.
        /// </summary>
        /// <remarks>
        /// The result must be disposed by the caller.
        /// </remarks>
        protected ICodeModule CodeModule()
        {
            var component = ProjectsProvider.Component(Module);
            return component?.CodeModule;
        }


        public virtual bool IsDirty
        {
            get
            {
                using (var codeModule = CodeModule())
                {
                    return codeModule == null || codeModule.Content() != Rewriter.GetText();
                }
            }
        }

        public ITokenStream TokenStream => Rewriter.TokenStream;

        public virtual void Rewrite()
        {
            if (!IsDirty)
            {
                return;
            }

            using (var codeModule = CodeModule())
            {
                codeModule.Clear();
                var newContent = Rewriter.GetText();
                codeModule.InsertLines(1, newContent);
            }
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
        }

        public void Remove(ITerminalNode target)
        {
            Rewriter.Delete(target.Symbol.TokenIndex);
        }

        public void Remove(IToken target)
        {
            Rewriter.Delete(target);
        }

        public void RemoveRange(int start, int stop)
        {
            Rewriter.Delete(start, stop);
        }

        public void Replace(Declaration target, string content)
        {
            Rewriter.Replace(target.Context.Start.TokenIndex, target.Context.Stop.TokenIndex, content);
        }

        public void Replace(ParserRuleContext target, string content)
        {
            Rewriter.Replace(target.Start.TokenIndex, target.Stop.TokenIndex, content);
        }

        public void Replace(IToken token, string content)
        {
            Rewriter.Replace(token, content);
        }

        public void Replace(ITerminalNode target, string content)
        {
            Rewriter.Replace(target.Symbol.TokenIndex, content);
        }

        public void Replace(Interval tokenInterval, string content)
        {
            Rewriter.Replace(tokenInterval.a, tokenInterval.b, content);
        }

        public void InsertBefore(int tokenIndex, string content)
        {
            Rewriter.InsertBefore(tokenIndex, content);
        }

        public void InsertAfter(int tokenIndex, string content)
        {
            Rewriter.InsertAfter(tokenIndex, content);
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