using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public abstract class DeclarationDeletionTargetBase : IComparable<DeclarationDeletionTargetBase>, IComparable
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly HashSet<Declaration> _targets;
        private readonly IModuleRewriter _rewriter;

        private List<Declaration> _allDeclarationsInListContext;
        
        public DeclarationDeletionTargetBase(IDeclarationFinderProvider declarationFinderProvider, Declaration target, IModuleRewriter rewriter)
        {
            if (target is null || declarationFinderProvider is null || rewriter is null)
            {
                throw new ArgumentNullException();
            }

            _declarationFinderProvider = declarationFinderProvider;
            _targets = new HashSet<Declaration>()
            {
                target
            };

            _rewriter = rewriter;
        }

        protected HashSet<Declaration> Targets => _targets;

        protected IDeclarationFinderProvider DeclarationFinderProvider => _declarationFinderProvider;

        public IModuleRewriter Rewriter => _rewriter;

        public virtual bool IsFullDelete => true;

        public void AddTargets(IEnumerable<Declaration> targets)
        {
            if (ListContext != null && targets.Any(t => GetListContext(t) != ListContext))
            {
                throw new ArgumentException("Attempted to add a delete target from a different list context");
            }

            foreach (var t in targets)
            {
                _targets.Add(t);
            }
        }

        public Declaration TargetProxy => _targets.First();

        public IReadOnlyCollection<Declaration> Declarations => _targets.ToList();

        //e.g., ModuleBodyElementContext, ModuleDeclarationElementContext, BlockContext
        public ParserRuleContext TargetContext { protected set; get; }

        public IReadOnlyList<Declaration> AllDeclarationsInListContext
        {
            get
            {
                if (_allDeclarationsInListContext is null)
                {
                    _allDeclarationsInListContext = ListContext != null
                        ? _declarationFinderProvider.DeclarationFinder.UserDeclarations(TargetProxy.DeclarationType)
                            .Where(d => GetListContext(d) == ListContext)
                            .Select(d => d)
                            .ToList()
                        : new List<Declaration>() { TargetProxy };
                }
                return _allDeclarationsInListContext;
            }
        }

        public IEnumerable<Declaration> RetainedDeclarations
            => AllDeclarationsInListContext.Except(_targets).ToList();

        public virtual VBAParser.EndOfStatementContext PrecedingEOSContext { set; get; }

        public VBAParser.EndOfStatementContext TargetEOSContext { protected set; get; }

        public VBAParser.EndOfStatementContext EOSContextToReplace => PrecedingEOSContext ?? TargetEOSContext;

        public virtual ParserRuleContext DeleteContext { protected set; get; }

        public ParserRuleContext ListContext { protected set; get; }

        public string ModifiedTargetEOSContent =>
            TargetEOSContext != null
                ? TargetEOSContext.CurrentContent(Rewriter)
                : string.Empty;

        private string CurrentPrecedingEOSContent =>
            PrecedingEOSContext != null
                ? PrecedingEOSContext.CurrentContent(Rewriter)
                : string.Empty;

        private static string RemoveStartingNewLines(string content)
            => string.Concat(content.SkipWhile(c => c == '\r' || c == '\n'));

        public virtual string BuildEOSReplacementContent()
        {
            string body;
            if (ModifiedTargetEOSContent.Contains(Tokens.CommentMarker))
            {
                if (CurrentPrecedingEOSContent.Contains(Tokens.CommentMarker))
                {
                    body = GetCurrentTextPriorToSeparationAndIndentation(PrecedingEOSContext, Rewriter)
                        + GetCurrentTextPriorToSeparationAndIndentation(TargetEOSContext, Rewriter);
                }
                else
                {
                    var eosContentPriorToSeparationAndIndentation = GetCurrentTextPriorToSeparationAndIndentation(TargetEOSContext, Rewriter);

                    body = GetCurrentTextPriorToSeparationAndIndentation(PrecedingEOSContext, Rewriter)
                        + PrecedingEOSSeparation
                        + RemoveStartingNewLines(eosContentPriorToSeparationAndIndentation);
                }

                return body + EOSSeparation + EOSIndentation;
            }

            body = GetCurrentTextPriorToSeparationAndIndentation(PrecedingEOSContext, Rewriter);

            return body + PrecedingEOSSeparation + EOSIndentation;
        }

        protected string PrecedingEOSSeparation => PrecedingEOSContext.GetSeparation();

        public string EOSSeparation => TargetEOSContext.GetSeparation();

        public string EOSIndentation => TargetEOSContext.GetIndentation();

        public virtual bool DeletionIncludesEOSContext { protected set;  get; } = true;

        public VBAParser.CommentContext GetDeclarationLogicalLineCommentContext() 
        {
            var individualNonEOFEOS = TargetEOSContext?.individualNonEOFEndOfStatement().FirstOrDefault();

            return individualNonEOFEOS?.GetDescendent<VBAParser.CommentContext>();
        }

        protected static ParserRuleContext GetListContext(Declaration target)
        {
            switch (target.DeclarationType)
            {
                case DeclarationType.Variable:
                    return target.Context.GetAncestor<VBAParser.VariableListStmtContext>();
                case DeclarationType.Constant:
                    return target.Context.GetAncestor<VBAParser.ConstStmtContext>();
                default:
                    return null;
            }
        }

        protected static string GetCurrentTextPriorToSeparationAndIndentation(VBAParser.EndOfStatementContext eosContext, IModuleRewriter rewriter)
        {
            var eosContent = eosContext.GetSeparationAndIndentationContent();
            var modifiedContent = eosContext.CurrentContent(rewriter);

            return modifiedContent.EndsWith(eosContent)
                ? modifiedContent.Substring(0, modifiedContent.Length - eosContent.Length)
                : modifiedContent;
        }

        public int CompareTo(DeclarationDeletionTargetBase other)
        {
            return other is null ? -1 : CompareTo(other);
        }

        public int CompareTo(object obj)
        {

            if (obj != null && !(obj is DeclarationDeletionTargetBase))
            {
                throw new ArgumentException("Object must be of type DeclarationDeletionTargetBase.");
            }

            var other = obj as DeclarationDeletionTargetBase;
            var thisFirstDeclaration = Declarations.OrderBy(d => d.Selection).FirstOrDefault();
            var otherFirstDeclaration = other.Declarations.OrderBy(d => d.Selection).FirstOrDefault();

            if (thisFirstDeclaration.Selection < otherFirstDeclaration.Selection)
            {
                return -1;
            }

            return thisFirstDeclaration.Selection == otherFirstDeclaration.Selection ? 0 : 1;
        }
    }
}
