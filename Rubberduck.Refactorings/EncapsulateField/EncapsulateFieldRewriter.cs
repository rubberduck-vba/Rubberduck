using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRewriter : IModuleRewriter
    {
        void SetVariableVisiblity(Declaration element, string visibilityToken);
        void Rename(Declaration element, string newName);
        void MakeImplicitDeclarationTypeExplicit(Declaration element);
        void InsertAtEndOfFile(string content);
        string GetText(int maxConsecutiveNewLines);
        void Replace(RewriteReplacePair pair);
    }

    public struct RewriteReplacePair
    {
        public RewriteReplacePair(string text, ParserRuleContext context)
        {
            Text = text;
            Context = context;
        }
        public string Text { private set; get; }
        public ParserRuleContext Context { private set; get; }
    }

    public class EncapsulateFieldRewriter : IEncapsulateFieldRewriter
    {
        private IModuleRewriter _rewriter;

        public EncapsulateFieldRewriter(IModuleRewriter rewriter)
        {
            _rewriter = rewriter;
        }

        public static IEncapsulateFieldRewriter CheckoutModuleRewriter(IRewriteSession rewriteSession, QualifiedModuleName qmn)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(qmn);
            return new EncapsulateFieldRewriter(rewriter);
        }

        public void InsertAtEndOfFile(string content)
        {
            if (content == string.Empty) { return; }

            _rewriter.InsertBefore(_rewriter.TokenStream.Size - 1, content);
        }

        public void SetVariableVisiblity(Declaration element, string visibility)
        {
            if (!element.IsVariable()) { return; }

            var variableStmtContext = element.Context.GetAncestor<VBAParser.VariableStmtContext>();
            var visibilityContext = variableStmtContext.GetChild<VBAParser.VisibilityContext>();

            if (visibilityContext != null)
            {
                _rewriter.Replace(visibilityContext, visibility);
                return;
            }
            _rewriter.InsertBefore(element.Context.Start.TokenIndex, $"{visibility} ");
        }

        public void Rename(Declaration element, string newName)
        {
            var identifierContext = element.Context.GetChild<VBAParser.IdentifierContext>();
            _rewriter.Replace(identifierContext, newName);
        }

        public void MakeImplicitDeclarationTypeExplicit(Declaration element)
        {
            if (!element.Context.TryGetChildContext<VBAParser.AsTypeClauseContext>(out _))
            {
                _rewriter.InsertAfter(element.Context.Stop.TokenIndex, $" {Tokens.As} {element.AsTypeName}");
            }
        }

        public string GetText(int maxConsecutiveNewLines)
        {
            var preview = GetText();
            var target = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines).ToList());
            var replacement = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines - 1).ToList());
            for (var counter = 1; counter < 10 && preview.Contains(target); counter++)
            {
                preview = preview.Replace(target, replacement);
            }
            return preview;
        }

        public bool IsDirty => _rewriter.IsDirty;

        public Selection? Selection { get => _rewriter.Selection; set => _rewriter.Selection = value; }

        public Selection? SelectionOffset { get => _rewriter.SelectionOffset; set => _rewriter.SelectionOffset = value; }

        public ITokenStream TokenStream => _rewriter.TokenStream;

        public string GetText(int startTokenIndex, int stopTokenIndex) => _rewriter.GetText(startTokenIndex, stopTokenIndex);

        public string GetText() => _rewriter.GetText();

        public void InsertAfter(int tokenIndex, string content) => _rewriter.InsertAfter(tokenIndex, content);

        public void InsertBefore(int tokenIndex, string content) => _rewriter.InsertBefore(tokenIndex, content);

        public void Remove(Declaration target) => _rewriter.Remove(target);

        public void Remove(ParserRuleContext target) => _rewriter.Remove(target);

        public void Remove(IToken target) => _rewriter.Remove(target);

        public void Remove(ITerminalNode target) => _rewriter.Remove(target);

        public void Remove(IParseTree target) => _rewriter.Remove(target);

        public void RemoveRange(int start, int stop) => _rewriter.RemoveRange(start, stop);

        public void Replace(Declaration target, string content) => _rewriter.Replace(target, content);

        public void Replace(RewriteReplacePair pair) => _rewriter.Replace(pair.Context, pair.Text);

        public void Replace(ParserRuleContext target, string content) => _rewriter.Replace(target, content);

        public void Replace(IToken token, string content) => _rewriter.Replace(token, content);

        public void Replace(ITerminalNode target, string content) => _rewriter.Replace(target, content);

        public void Replace(IParseTree target, string content) => _rewriter.Replace(target, content);

        public void Replace(Interval tokenInterval, string content) => _rewriter.Replace(tokenInterval, content);
    }

}
