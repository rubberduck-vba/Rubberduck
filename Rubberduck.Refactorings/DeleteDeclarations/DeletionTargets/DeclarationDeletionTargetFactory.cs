using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeleteDeclarations;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public class DeclarationDeletionTargetFactory : IDeclarationDeletionTargetFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public DeclarationDeletionTargetFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IEnumerable<IDeclarationDeletionTarget> CreateMany(IEnumerable<Declaration> declarations, IRewriteSession rewriteSession)
        {
            return declarations.Select(d => Create(d, rewriteSession));
        }

        public IDeclarationDeletionTarget Create(Declaration declaration, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);

            if (declaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return declaration.DeclarationType.HasFlag(DeclarationType.Property)
                    ? new PropertyMemberDeletionTarget(_declarationFinderProvider, declaration, rewriter)
                    : new ModuleElementDeletionTarget(_declarationFinderProvider, declaration, rewriter);
            }

            switch (declaration.DeclarationType)
            {
                case DeclarationType.UserDefinedTypeMember:
                    return new UdtMemberDeletionTarget(_declarationFinderProvider, declaration, rewriter);

                case DeclarationType.EnumerationMember:
                    return new EnumMemberDeletionTarget(_declarationFinderProvider, declaration, rewriter);

                case DeclarationType.Variable:
                    return declaration.ParentDeclaration is ModuleDeclaration
                        ? new ModuleElementDeletionTarget(_declarationFinderProvider, declaration, rewriter)
                        : new ProcedureLocalDeletionTarget<VBAParser.VariableListStmtContext>(_declarationFinderProvider, declaration, rewriter) as IDeclarationDeletionTarget;

                case DeclarationType.Constant:
                    return declaration.ParentDeclaration is ModuleDeclaration
                        ? new ModuleElementDeletionTarget(_declarationFinderProvider, declaration, rewriter)
                        : new ProcedureLocalDeletionTarget<VBAParser.ConstStmtContext>(_declarationFinderProvider, declaration, rewriter) as IDeclarationDeletionTarget;

                case DeclarationType.Enumeration:
                case DeclarationType.UserDefinedType:
                    return new ModuleElementDeletionTarget(_declarationFinderProvider, declaration, rewriter);

                case DeclarationType.LineLabel:
                    return new LineLabelDeletionTarget(_declarationFinderProvider, declaration, rewriter);
                default:
                    throw new ArgumentException();
            }
        }
    }
}
