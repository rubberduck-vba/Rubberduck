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
    //TODO: Could use this as a base class for ModuleFieldDT, ModuleConstantDT, MemberDT, UDTDT, EnumDT
    internal class ModuleElementDeletionTarget : DeleteDeclarationTarget, IModuleElementDeletionTarget
    {
        public ModuleElementDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target)
            : base(declarationFinderProvider, target)
        {
            if (target.Context.TryGetAncestor<VBAParser.ModuleDeclarationsElementContext>(out var mde))
            {
                _targetContext = mde;
            }
            else if (target.Context.TryGetAncestor<VBAParser.ModuleBodyElementContext>(out var mbe))
            {
                _targetContext = mbe;
            }
            
            if (target.DeclarationType == DeclarationType.Variable
                || target.DeclarationType == DeclarationType.Constant)
            {
                _listContext = GetListContext(target);
            }

            _deleteContext = target.DeclarationType.HasFlag(DeclarationType.Member)
                ? target.Context.GetAncestor<VBAParser.ModuleBodyElementContext>()
                : target.Context.GetAncestor<VBAParser.ModuleDeclarationsElementContext>() as ParserRuleContext;

            _eosContext = GetFollowingEndOfStatementContext(_deleteContext);
        }

        public override bool IsFullDelete 
            => DeclarationType != DeclarationType.Variable && TargetProxy.DeclarationType != DeclarationType.Constant
                || _allDeclarationsInList.Intersect(_targets).Count() == _allDeclarationsInList.Count;
        
        public void SetPrecedingEOSContext(VBAParser.EndOfStatementContext eos)
        {
            if (TargetProxy.ParentDeclaration is ModuleDeclaration)
            {
                _precedingEOSContext = eos;
            }
        }
    }
}
