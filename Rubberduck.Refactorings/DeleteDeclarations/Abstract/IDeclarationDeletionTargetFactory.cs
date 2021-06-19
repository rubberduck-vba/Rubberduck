using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeleteDeclarations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings
{
    public interface IDeclarationDeletionTargetFactory
    {
        IDeclarationDeletionTarget Create(Declaration declaration);
        IEnumerable<IDeclarationDeletionTarget> CreateMany(IEnumerable<Declaration> declarations);
    }
    public class DeclarationDeletionTargetFactory : IDeclarationDeletionTargetFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public DeclarationDeletionTargetFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        //TODO: yield?
        public IEnumerable<IDeclarationDeletionTarget> CreateMany(IEnumerable<Declaration> declarations)
        {
            return declarations.Select(d => Create(d));
        }

        public IDeclarationDeletionTarget Create(Declaration declaration)
        {
            if (declaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return new ModuleElementDeletionTarget(_declarationFinderProvider, declaration);
            }

            switch (declaration.DeclarationType)
            {
                case DeclarationType.UserDefinedTypeMember:
                    return new UdtMemberDeletionTarget(_declarationFinderProvider, declaration);
                case DeclarationType.EnumerationMember:
                    return new EnumMemberDeletionTarget(_declarationFinderProvider, declaration);
                case DeclarationType.Variable:
                    return declaration.ParentDeclaration is ModuleDeclaration
                        ? new ModuleElementDeletionTarget(_declarationFinderProvider, declaration)
                        : new ProcedureLocalDeletionTarget<VBAParser.VariableListStmtContext>(_declarationFinderProvider, declaration) as IDeclarationDeletionTarget;
                case DeclarationType.Constant:
                    return declaration.ParentDeclaration is ModuleDeclaration
                        ? new ModuleElementDeletionTarget(_declarationFinderProvider, declaration)
                        : new ProcedureLocalDeletionTarget<VBAParser.ConstStmtContext>(_declarationFinderProvider, declaration) as IDeclarationDeletionTarget;
                case DeclarationType.Enumeration:
                case DeclarationType.UserDefinedType:
                    return new ModuleElementDeletionTarget(_declarationFinderProvider, declaration);
                case DeclarationType.LineLabel:
                    return new LineLabelDeletionTarget(_declarationFinderProvider, declaration);
                default:
                    throw new ArgumentException();
            }
        }
    }
}
