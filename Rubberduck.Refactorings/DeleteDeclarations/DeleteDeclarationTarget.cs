using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteDeclarationTarget
    {
        protected readonly IDeclarationFinderProvider _declarationFinderProvider;

        protected readonly HashSet<Declaration> _targets; //part of the same list context

        protected List<Declaration> _allDeclarationsInList;

        protected ParserRuleContext _deleteContext;
        protected VBAParser.EndOfStatementContext _eosContext;
        protected ParserRuleContext _listContext;
        protected ParserRuleContext _targetContext; //e.g., ModuleBodyElementContext, ModuleDeclarationElementContext, BlockContext

        protected VBAParser.EndOfStatementContext _precedingEOSContext;

        public DeleteDeclarationTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target)
        {
            if (target is null || declarationFinderProvider is null)
            {
                throw new ArgumentNullException();
            }

            _declarationFinderProvider = declarationFinderProvider;
            _targets = new HashSet<Declaration>()
            {
                target
            };
        }

        public virtual bool IsFullDelete => true;

        public void AddTargets(IEnumerable<Declaration> targets)
        {
            //TODO: Check for same ListContext?
            foreach (var t in targets)
            {
                _targets.Add(t);
            }
        }

        public DeclarationType DeclarationType => TargetProxy.DeclarationType;

        public Declaration TargetProxy => _targets.First();

        public ParserRuleContext TargetContext => _targetContext;

        public List<Declaration> AllDeclarationsInListContext
        {
            get
            {
                if (_allDeclarationsInList is null)
                {
                    _allDeclarationsInList = _listContext != null
                        ? _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType)
                            .Where(d => GetListContext(d) == _listContext)
                            .Select(d => d)
                            .ToList()
                        : new List<Declaration>() { TargetProxy };
                }
                return _allDeclarationsInList;
            }
        }

        public IEnumerable<Declaration> RetainedDeclarations
            => _allDeclarationsInList.Except(_targets).ToList();

        public virtual VBAParser.EndOfStatementContext PrecedingEOSContext => _precedingEOSContext;

        public VBAParser.EndOfStatementContext EndOfStatementContext => _eosContext;

        public virtual ParserRuleContext DeleteContext => _deleteContext;

        public ParserRuleContext ListContext => _listContext;

        protected static VBAParser.EndOfStatementContext GetFollowingEndOfStatementContext(ParserRuleContext context)
        {
            context.TryGetFollowingContext(out VBAParser.EndOfStatementContext eos); 
            return eos;
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
    }
}
