using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings
{
    public interface IModuleElementDeletionTarget : IDeclarationDeletionTarget    
    {
        void SetPrecedingEOSContext(VBAParser.EndOfStatementContext eos);
    }
    public interface IEnumMemberDeletionTarget : IDeclarationDeletionTarget
    {

    }
    public interface IUdtMemberDeletionTarget : IDeclarationDeletionTarget
    {

    }
    public interface IProcedureLocalDeletionTarget : IDeclarationDeletionTarget
    {
        bool DeleteAssociatedLabel { set; get; }
    }
    public interface ILineLabelDeletionTarget : IDeclarationDeletionTarget
    {

    }
    public interface IDeclarationDeletionTarget
    {
        bool IsFullDelete { get; }

        void AddTargets(IEnumerable<Declaration> targets);

        DeclarationType DeclarationType { get; }

        Declaration TargetProxy { get; }

        List<Declaration> AllDeclarationsInListContext { get; }

        IEnumerable<Declaration> RetainedDeclarations { get; }

        VBAParser.EndOfStatementContext PrecedingEOSContext { get; }

        VBAParser.EndOfStatementContext EndOfStatementContext { get; }

        ParserRuleContext DeleteContext { get; }

        //TODO: Only applies to Variable and Constants
        ParserRuleContext ListContext { get; }

        ParserRuleContext TargetContext { get; }
    }
}
