using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    /// <summary>
    /// Removes 1 to n Declarations and optionally applies Indenter settings to each module containing deleted Declaration(s).  
    /// Warning: If references to the removed Declarations are not removed/modified by other code, 
    /// this refactoring action will generate uncompilable code.
    /// </summary>
    public class DeleteDeclarationsRefactoringAction : CodeOnlyRefactoringActionBase<DeleteDeclarationsModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IIndenter _indenter;

        private static readonly string _lineExtensionExpression = $" _{Environment.NewLine}";

        private List<DeclarationType> _supportedDeclarationTypes;

        public DeleteDeclarationsRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager, IIndenter indenter)
            : base(rewritingManager) 
        {
            _declarationFinderProvider = declarationFinderProvider;
            _indenter = indenter;
            _supportedDeclarationTypes = new List<DeclarationType>() 
            {
                DeclarationType.Variable,
                DeclarationType.Constant,
                DeclarationType.Function,
                DeclarationType.Procedure,
                DeclarationType.PropertyGet,
                DeclarationType.PropertyLet,
                DeclarationType.PropertySet,
                DeclarationType.UserDefinedType,
                DeclarationType.UserDefinedTypeMember,
                DeclarationType.Enumeration,
                DeclarationType.EnumerationMember,
                DeclarationType.LineLabel
            };
        }

        public override void Refactor(DeleteDeclarationsModel model, IRewriteSession rewriteSession)
        {
            var targetsByDeclarationType = new Dictionary<DeclarationType, ILookup<QualifiedModuleName, Declaration>>();
            
            foreach (var key in _supportedDeclarationTypes)
            {
                targetsByDeclarationType[key] = OrganizeDeclarationsByQMN(model, key);
            }

            RefactorVariablesAndConstants(targetsByDeclarationType, _declarationFinderProvider, rewriteSession);

            RefactorRemainingTypes(targetsByDeclarationType, rewriteSession);

            if (model.IndentModifiedModules)
            {
                var qmnsToIndent = new List<QualifiedModuleName>();
                foreach (var decType in _supportedDeclarationTypes.Where(t => t.HasFlag(DeclarationType.Member)))
                {
                    qmnsToIndent.AddRange(targetsByDeclarationType[decType].Select(d => d.Key));
                }

                var startTokenIndex = 0;
                foreach (var qmn in qmnsToIndent.Distinct())
                {
                    var rewriter = rewriteSession.CheckOutModuleRewriter(qmn);
                    var stopTokenIndex = rewriter.TokenStream.Size - 1;

                    var contentToModify = rewriter.GetText(startTokenIndex, stopTokenIndex);
                    var lines = _indenter.Indent(contentToModify);

                    var formattedContent = string.Join(Environment.NewLine, lines);

                    rewriter.Replace(new Interval(startTokenIndex, stopTokenIndex), formattedContent);
                }
            }
        }

        private static void RefactorVariablesAndConstants(Dictionary<DeclarationType, ILookup<QualifiedModuleName, Declaration>> targetsByDeclarationType, IDeclarationFinderProvider declarationFinderProvider, IRewriteSession rewriteSession)
        {
            if (targetsByDeclarationType[DeclarationType.Variable].Any())
            {
                foreach (var variablesGrouping in targetsByDeclarationType[DeclarationType.Variable])
                {
                    var rewriter = rewriteSession.CheckOutModuleRewriter(variablesGrouping.Key);
                    RemoveVariablesOrConstants<VBAParser.VariableListStmtContext, VBAParser.VariableSubStmtContext>(variablesGrouping, declarationFinderProvider, rewriter);
                }
            }

            if (targetsByDeclarationType[DeclarationType.Constant].Any())
            {
                foreach (var constantsGrouping in targetsByDeclarationType[DeclarationType.Constant])
                {
                    var rewriter = rewriteSession.CheckOutModuleRewriter(constantsGrouping.Key);
                    RemoveVariablesOrConstants<VBAParser.ConstStmtContext, VBAParser.ConstSubStmtContext>(constantsGrouping, declarationFinderProvider, rewriter);
                }
            }
        }

        private static void RemoveVariablesOrConstants<TListContext, TSubStmtContext>(IEnumerable<Declaration> toRemove, IDeclarationFinderProvider declarationFinderProvider, IModuleRewriter rewriter) where TListContext : ParserRuleContext where TSubStmtContext : ParserRuleContext
        {
            if (!toRemove.Any())
            {
                return;
            }

            var targetsToDeleteByListContext = toRemove.Distinct()
                .GroupBy(f => f.Context.GetAncestor<TListContext>());

            foreach (var targetsToDelete in targetsToDeleteByListContext)
            {
                RewriteVariableOrConstantListContextDeclarations(targetsToDelete, declarationFinderProvider, rewriter);
            }
        }

        private static void RewriteVariableOrConstantListContextDeclarations<TListContext>(IGrouping<TListContext, Declaration> targetsToDelete,
            IDeclarationFinderProvider declarationFinderProvider,
            IModuleRewriter rewriter) where TListContext : ParserRuleContext
        {
            var listContext = targetsToDelete.Key;
            var listItemProxy = targetsToDelete.First();

            var declarationsToRetain = declarationFinderProvider.DeclarationFinder.UserDeclarations(listItemProxy.DeclarationType)
                .Where(d => d.Context.GetAncestor<TListContext>() == targetsToDelete.Key)
                .Except(targetsToDelete)
                .ToList();

            if (declarationsToRetain.Any())
            {
                //Delete a subset of the the declaration list
                var retainedDeclarationsExpression = listContext.GetText().Contains(_lineExtensionExpression)
                    ? $"{BuildDeclarationsExpressionWithLineContinuations(targetsToDelete, declarationsToRetain)}"
                    : $"{string.Join(", ", declarationsToRetain.Select(d => d.Context.GetText()))}";

                var replacementExpression = $"{GetAccessibiltyToken(listItemProxy)} {retainedDeclarationsExpression}";

                rewriter.Replace(listContext.Parent, replacementExpression);
                return;
            }

            //Delete the entire declaration list
            if (listItemProxy.Context.TryGetAncestor<VBAParser.ModuleDeclarationsElementContext>(out var mdeContext))
            {
                rewriter.Remove(mdeContext);
                ModifyEndOfStatementContext(mdeContext, rewriter);
            }
            else if (listItemProxy.Context.TryGetAncestor<VBAParser.BlockStmtContext>(out var blockStmtContext))
            {
                var forceRetentionOfEndOfStatementContent = false;
                ParserRuleContext contextToRemove = blockStmtContext;

                if (blockStmtContext.TryGetChildContext<VBAParser.StatementLabelDefinitionContext>(out var labelDefinitionContext))
                {
                    //The blockStmtContext contains a label which should be preserved
                    if (!blockStmtContext.TryGetFollowingContext<VBAParser.EndOfStatementContext>(out var eosContext))
                    {
                        throw new ArgumentException("Unable to get expected VBAParser.EndOfStatementContext");
                    }

                    forceRetentionOfEndOfStatementContent = true;
                    contextToRemove = listContext.Parent as ParserRuleContext;
                }

                rewriter.Remove(contextToRemove);
                ModifyEndOfStatementContext(blockStmtContext, rewriter, forceRetentionOfEndOfStatementContent);
            }
            else
            {
                throw new ArgumentException("Unable to get Variable/Constant ancestor 'ModuleDeclarationsElementContext' or 'BlockStmtContext'");
            }
        }

        private static void RefactorRemainingTypes(Dictionary<DeclarationType, ILookup<QualifiedModuleName, Declaration>> targetsByDeclarationType, IRewriteSession rewriteSession)
        {
            var declarationTypes = targetsByDeclarationType.Keys.Except(new DeclarationType[] { DeclarationType.Variable, DeclarationType.Constant }).ToList();
            foreach (var decType in declarationTypes)
            {
                if (!targetsByDeclarationType[decType].Any())
                {
                    continue;
                }

                RemoveGroupings(rewriteSession, targetsByDeclarationType[decType]);
            }
        }

        private static void RemoveGroupings(IRewriteSession rewriteSession, ILookup<QualifiedModuleName, Declaration> declarationsByQMN)
        {
            foreach (var declarationGroup in declarationsByQMN)
            {
                if (!declarationGroup.Any())
                {
                    return;
                }

                var rewriter = rewriteSession.CheckOutModuleRewriter(declarationGroup.Key);
                foreach (var dec in declarationGroup)
                {
                    rewriter.Remove(dec);
                    ModifyEndOfStatementContext(dec.Context, rewriter);
                }
            }
        }

        private static VBAParser.EndOfStatementContext GetEndOfStmtContext(ParserRuleContext context)
        {
            var  eosPredecessor = context.Parent as ParserRuleContext;

            switch (context)
            {
                case VBAParser.EndOfStatementContext eosCtxt:
                    return eosCtxt;
                case VBAParser.UdtMemberContext _:
                case VBAParser.ModuleDeclarationsElementContext _:
                case VBAParser.StatementLabelDefinitionContext _:
                case VBAParser.BlockStmtContext _:
                    eosPredecessor = context;
                    break;
                case VBAParser.EnumerationStmt_ConstantContext _:
                    return null;
                case VBAParser.IdentifierStatementLabelContext _:
                    //If the label to remove is the only content in the block statement, then remove the end of statement context as well
                    if (context.TryGetAncestor<VBAParser.BlockStmtContext>(out var blockStmt)
                        && string.Compare(context.GetText(), blockStmt.GetText(), StringComparison.InvariantCulture) == 0
                        && blockStmt.TryGetFollowingContext<VBAParser.EndOfStatementContext>(out var eosCt))
                    {
                        return eosCt;
                    }
                    return null;
            }

            if (eosPredecessor.TryGetFollowingContext<VBAParser.EndOfStatementContext>(out var eosContext))
            {
                return eosContext;
            }
            throw new ArgumentException("Unable to get expected VBAParser.EndOfStatementContext");
        }

        private static void ModifyEndOfStatementContext(ParserRuleContext prContext, IModuleRewriter rewriter, bool forceRetentionOfEndOfStatementContent = false)
        {
            var eosContext = GetEndOfStmtContext(prContext);
            if (eosContext is null)
            {
                return;
            }

            var eosContent = eosContext.GetText();

            var replacement = eosContent.Contains("'")
                ? RemoveCommentIfOnSameLogicalLine(eosContent)
                : string.Empty;

            if (forceRetentionOfEndOfStatementContent 
                && replacement.Length == 0
                //If the EndOfStatementContext did not originally have a newLine, (e.g., ": ") do not inject a newLine
                && eosContent.Contains(Environment.NewLine)) 
            {
                replacement = eosContent;
            }
            rewriter.Replace(eosContext, replacement);
        }

        private static string RemoveCommentIfOnSameLogicalLine(string eosContent)
        {
            var replacement = eosContent;

            //Remove line extension newLines
            if (eosContent.Contains(_lineExtensionExpression))
            {
                replacement = eosContent.Replace(_lineExtensionExpression, " _");
            }

            var indexOfFirstNewLine = replacement.IndexOf(Environment.NewLine);

            if (indexOfFirstNewLine >= 0)
            {
                replacement = replacement.Substring(indexOfFirstNewLine + Environment.NewLine.Length);
            }

            //Restore line extensions and associated newLines
            if (eosContent.Contains(_lineExtensionExpression))
            {
                replacement = replacement.Replace(" _", _lineExtensionExpression);
            }
            return replacement;
        }

        private static string GetAccessibiltyToken(Declaration listPrototype)
        {
            bool isModuleScope = listPrototype.ParentDeclaration is ModuleDeclaration;

            var accessToken = listPrototype.Accessibility == Accessibility.Implicit
                ? Tokens.Private
                : $"{listPrototype.Accessibility}";

            if (listPrototype.DeclarationType == DeclarationType.Constant)
            {
                accessToken = isModuleScope ? $"{accessToken} {Tokens.Const}" : Tokens.Const;
            }

            if (listPrototype.DeclarationType == DeclarationType.Variable && !isModuleScope)
            {
                accessToken = Tokens.Dim;
            }
            return accessToken;
        }

        private static string BuildDeclarationsExpressionWithLineContinuations<T>(IGrouping<T, Declaration> targetsToDelete, List<Declaration> toRetain) where T : ParserRuleContext
        {
            var elementsByLineContinuation = targetsToDelete.Key.GetText().Split(new string[] { _lineExtensionExpression }, StringSplitOptions.None);

            if (elementsByLineContinuation.Count() == 1)
            {
                throw new ArgumentException("'targetsToDelete' parameter does not contain line extension(s)");
            }

            var expr = new StringBuilder();
            foreach (var element in elementsByLineContinuation)
            {
                var idContexts = toRetain.Where(r => element.Contains(r.Context.GetText())).Select(d => d);
                foreach (var ctxt in idContexts)
                {
                    var indent = string.Concat(element.TakeWhile(e => e == ' '));

                    expr = expr.Length == 0
                        ? expr.Append(ctxt.Context.GetText())
                        : expr.Append($",{_lineExtensionExpression}{indent}{ctxt.Context.GetText()}");
                }
            }
            return expr.ToString();
        }

        private static ILookup<QualifiedModuleName, Declaration> OrganizeDeclarationsByQMN(DeleteDeclarationsModel model, DeclarationType declarationType)
            => model.Targets.Where(d => d.DeclarationType == declarationType).ToLookup(v => v.QualifiedModuleName);
    }
}
