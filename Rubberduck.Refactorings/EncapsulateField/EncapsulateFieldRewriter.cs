using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRewriter : IModuleRewriter
    {
        //void RemoveDeclarations(IEnumerable<Declaration> allDeclarationsToRemove, IEnumerable<Declaration> allModuleDeclarations);
        void InsertNewContent(int? codeSectionStartIndex, IEncapsulateFieldNewContentProvider newContent);
        void ReplaceDescendentContext<T>(Declaration member, string content) where T : ParserRuleContext;
        void InsertBeforeDescendentContext<T>(Declaration member, string content) where T : ParserRuleContext;
        void SetMemberVisibility(Declaration member, string visibilityToken);
        void SetVariableVisiblity(Declaration element, string visibilityToken);
        void Rename(Declaration element, string newName);
        void MakeImplicitDeclarationTypeExplicit(Declaration element);
        void RemoveMemberAccess(IEnumerable<IdentifierReference> memberReferences);
        void RemoveMemberAccess(IdentifierReference idRef);
        void RemoveWithMemberAccess(IEnumerable<IdentifierReference> idReferences);
        void InsertAtEndOfFile(string content);
        string GetModifiedText(Declaration declaration);
    }

    public class EncapsulateFieldRewriter : IEncapsulateFieldRewriter
    {
        public static IEncapsulateFieldRewriter CheckoutModuleRewriter(IRewriteSession rewriteSession, QualifiedModuleName qmn)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(qmn);
            return new EncapsulateFieldRewriter(rewriter);
        } 

        private IModuleRewriter _rewriter;

        public EncapsulateFieldRewriter(IModuleRewriter rewriter)
        {
            _rewriter = rewriter;
        }

        public void InsertNewContent(int? codeSectionStartIndex, IEncapsulateFieldNewContentProvider newContent)
        {
            if (codeSectionStartIndex.HasValue && newContent.HasNewContent)
            {
                _rewriter.InsertBefore(codeSectionStartIndex.Value, $"{Environment.NewLine}{newContent.AsSingleTextBlock}");
            }
            else
            {
                InsertAtEndOfFile($"{Environment.NewLine}{newContent.AsSingleTextBlock}");
            }
        }

        public void ReplaceDescendentContext<T>(Declaration member, string content) where T : ParserRuleContext
        {
            var descendentCtxt = member.Context.GetDescendent<T>();
            if (descendentCtxt != null)
            {
                _rewriter.Replace(descendentCtxt, content);
            }
        }

        public void InsertBeforeDescendentContext<T>(Declaration member, string content) where T : ParserRuleContext
        {
            var descendentCtxt = member.Context.GetDescendent<T>();
            if (descendentCtxt != null)
            {
                _rewriter.InsertBefore(descendentCtxt.Start.TokenIndex, content);
            }
        }

        //public void RemoveDeclarations(IEnumerable<Declaration> allDeclarationsToRemove, IEnumerable<Declaration> allModuleDeclarations)
        //{
        //    if (allDeclarationsToRemove.Count() == 0)
        //    {
        //        return;
        //    }
        //    var original = _rewriter.GetText();

        //    var declaredInLists = allDeclarationsToRemove.Where(declaration =>
        //        declaration.Context.Parent is VBAParser.VariableListStmtContext
        //            || declaration.Context.Parent is VBAParser.ConstStmtContext);

        //    RemoveMany(_rewriter, allDeclarationsToRemove.Except(declaredInLists));

        //    //Handle special cases where the declarations to remove 
        //    //are/can be declared within a declarationlist context
        //    if (declaredInLists.Any())
        //    {
        //        var lookupCtxtToDeclarationListRemovals = declaredInLists.ToLookup(dec => dec.Context.Parent as ParserRuleContext);

        //        foreach (var declarationsToRemoveFromList in lookupCtxtToDeclarationListRemovals)
        //        {
        //            var allDeclarationsInList = allModuleDeclarations
        //               .Where(dec => dec.Context?.Parent == declarationsToRemoveFromList.Key);

        //            RemoveDeclarationListContent(declarationsToRemoveFromList, allDeclarationsInList);
        //        }
        //        var removeListyResult = _rewriter.GetText();
        //    }
        //}

        //The issues described within this method seems to have been fixed
        //private void RemoveDeclarationListContent(IGrouping<ParserRuleContext, Declaration> toRemoveFromDeclarationList, IEnumerable<Declaration> allDeclarationsInDeclarationList)
        //{
        //    //Remove the entire list
        //    if (toRemoveFromDeclarationList.Count() == allDeclarationsInDeclarationList.Count())
        //    {
        //        var parentContext = toRemoveFromDeclarationList.First().Context.Parent;
        //        if (parentContext is VBAParser.ConstStmtContext)
        //        {
        //            _rewriter.Remove(parentContext);
        //        }
        //        else
        //        {
        //            _rewriter.Remove(parentContext.Parent);
        //        }
        //        return;
        //    }

        //    //A subset of the declarations in the list are to be removed
        //    //1. Remove declarations individually
        //    //2. Handle special case described below
        //    RemoveMany(_rewriter, toRemoveFromDeclarationList);

        //    //Special case:
        //    //If there are 'n' declarations in a list (where 'n' >= 3) and we are removing 2 to n-1 of
        //    //the LAST declarations, calling 'rewriter.Remove' on each declaration leaves 
        //    //a trailing comma on the last RETAINED declaration.
        //    if (toRemoveFromDeclarationList.Count() >= 2 && allDeclarationsInDeclarationList.Count() >= 3)
        //    {
        //        var reversedDeclarationListElements = allDeclarationsInDeclarationList.OrderByDescending(tr => tr.Selection);
        //        var removedFromEndOfList = reversedDeclarationListElements.TakeWhile(rd => toRemoveFromDeclarationList.Contains(rd));
        //        if (removedFromEndOfList.Count() >= 2)
        //        {
        //            var lastRetainedDeclaration = reversedDeclarationListElements.ElementAt(removedFromEndOfList.Count());
        //            var tokenStart = lastRetainedDeclaration.Context.Stop.TokenIndex + 1;
        //            var tokenStop = removedFromEndOfList.Last().Context.Start.TokenIndex - 1;
        //            _rewriter.RemoveRange(tokenStart, tokenStop);
        //        }
        //    }
        //}

        private static void RemoveMany(IModuleRewriter rewriter, IEnumerable<Declaration> declarations)
        {
            foreach (var dec in declarations)
            {
                rewriter.Remove(dec);
            }
        }

        public void InsertAtEndOfFile(string content)
        {
            if (content == string.Empty)
            {
                return;
            }
            _rewriter.InsertBefore(_rewriter.TokenStream.Size - 1, content);
        }

        public string  GetModifiedText(Declaration declaration)
        {
            return _rewriter.GetText(declaration.Context.Start.TokenIndex, declaration.Context.Stop.TokenIndex);
        }

        public void SetMemberVisibility(Declaration element, string visibility)
        {
            if (!element.IsMember()) { return; }

            var visibilityContext = element.Context.GetChild<VBAParser.VisibilityContext>();
            if (visibilityContext != null)
            {
                _rewriter.Replace(visibilityContext, visibility);
                return;
            }
            _rewriter.InsertBefore(element.Context.Start.TokenIndex, $"{visibility} ");
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

        public void RemoveMemberAccess(IEnumerable<IdentifierReference> memberReferences)
        {
            var memberAccessExprContexts = memberReferences
                .Where(rf => rf.Context.Parent is VBAParser.MemberAccessExprContext);

            foreach (var context in memberAccessExprContexts)
            {
                RemoveMemberAccess(context);
            }
        }

        public void RemoveWithMemberAccess(IEnumerable<IdentifierReference> references)
        {
            foreach (var withMemberAccessExprContext in references.Where(rf => rf.Context.Parent is VBAParser.WithMemberAccessExprContext).Select(rf => rf.Context.Parent as VBAParser.WithMemberAccessExprContext))
            {
                RemoveRange(withMemberAccessExprContext.Start.TokenIndex, withMemberAccessExprContext.Start.TokenIndex);
            }
        }

        public void RemoveMemberAccess(IdentifierReference idRef)
        {
            if (idRef.Context.Parent is VBAParser.MemberAccessExprContext maec)
            {
                Debug.Assert(maec.ChildCount == 3, "MemberAccessExprContext child contexts does not equal 3");
                Replace(maec, maec.children[2].GetText());
            }
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

        public void Replace(ParserRuleContext target, string content) => _rewriter.Replace(target, content);

        public void Replace(IToken token, string content) => _rewriter.Replace(token, content);

        public void Replace(ITerminalNode target, string content) => _rewriter.Replace(target, content);

        public void Replace(IParseTree target, string content) => _rewriter.Replace(target, content);

        public void Replace(Interval tokenInterval, string content) => _rewriter.Replace(tokenInterval, content);
    }
}
