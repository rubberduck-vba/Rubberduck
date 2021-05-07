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
    internal class UdtMemberDeletionTarget : DeleteDeclarationTarget, IUdtMemberDeletionTarget
    {
        public UdtMemberDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target)
            : base(declarationFinderProvider, target) 
        {
            _targetContext = target.Context;
            _deleteContext = target.Context;
            
            _listContext = target.Context.GetAncestor<VBAParser.UdtMemberListContext>();
            
            _precedingEOSContext = _deleteContext == _listContext.children.First()
                ? TargetProxy.Context.GetAncestor<VBAParser.UdtDeclarationContext>()
                    .GetChild<VBAParser.EndOfStatementContext>()
                : _listContext.children.TakeWhile(ch => ch != TargetProxy.Context).Last() as VBAParser.EndOfStatementContext;
            
            _eosContext = GetFollowingEndOfStatementContext(_deleteContext);
        }
    }
}

