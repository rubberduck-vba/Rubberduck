using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.PostProcessing.RewriterInfo;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.PostProcessing
{
    public class ModuleRewriter : IModuleRewriter
    {
        private readonly ICodeModule _module;
        private readonly TokenStreamRewriter _rewriter;

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

        private static readonly IDictionary<DeclarationType, IRewriterInfoFinder> Finders =
            new Dictionary<DeclarationType, IRewriterInfoFinder>
            {
                {DeclarationType.Variable, new VariableRewriterInfoFinder()},
                {DeclarationType.Constant, new ConstantRewriterInfoFinder()},
                {DeclarationType.Parameter, new ParameterRewriterInfoFinder()},
            };

        public void Remove(Declaration target)
        {
            IRewriterInfoFinder finder;
            var info = Finders.TryGetValue(target.DeclarationType, out finder) 
                ? finder.GetRewriterInfo(target.Context, target) 
                : new DefaultRewriterInfoFinder().GetRewriterInfo(target.Context, target);            

            if (info.Equals(RewriterInfo.RewriterInfo.None)) { return; }

            _rewriter.Delete(info.StartTokenIndex, info.StopTokenIndex);

            var references = target.References.Where(r => r.QualifiedModuleName.Component.CodeModule == _module);
            foreach (var reference in references)
            {
                Remove(reference);
            }
        }

        public void Remove(IdentifierReference target)
        {
            _rewriter.Delete(target.Context.Start.TokenIndex, target.Context.Stop.TokenIndex);
        }

        public void Remove(ParserRuleContext target)
        {
            _rewriter.Delete(target.Start.TokenIndex, target.Stop.TokenIndex);
        }

        public void Replace(Declaration target, string content)
        {
            throw new System.NotImplementedException();
        }

        public void Replace(IdentifierReference target, string content)
        {
            throw new System.NotImplementedException();
        }

        public void Replace(ParserRuleContext target, string content)
        {
            throw new System.NotImplementedException();
        }

        public void Replace(IToken token, string content)
        {
            throw new System.NotImplementedException();
        }

        public void Rename(Declaration target, string identifier)
        {
            throw new System.NotImplementedException();
        }

        public void Insert(string content, int line = 1, int column = 1)
        {
            throw new System.NotImplementedException();
        }

        public void AppendToDeclarations(string content)
        {
            var line = _module.CountOfDeclarationLines + 1;
            Insert(content, line);
        }

        public void InsertAtIndex(string content, int tokenIndex)
        {
            throw new System.NotImplementedException();
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