using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
namespace Rubberduck.Refactorings.CreateUDTMember
{
    public class CreateUDTMemberRefactoringAction : CodeOnlyRefactoringActionBase<CreateUDTMemberModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ICodeBuilder _codeBuilder;

        public CreateUDTMemberRefactoringAction(IDeclarationFinderProvider declarationFinderProvider,IRewritingManager rewritingManager, ICodeBuilder codeBuilder)
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _codeBuilder = codeBuilder;
        }

        public override void Refactor(CreateUDTMemberModel model, IRewriteSession rewriteSession)
        {
            if (model.UserDefinedTypeTargets.Any( udt => !(udt.Context is VBAParser.UdtDeclarationContext)))
            {
                throw new ArgumentException();
            }

            foreach (var udt in model.UserDefinedTypeTargets)
            {
                InsertNewMembersBlock(BuildNewMembersBlock(udt, model[udt]),
                    GetInsertionIndex(udt.Context as VBAParser.UdtDeclarationContext),
                    rewriteSession.CheckOutModuleRewriter(udt.QualifiedModuleName));
            }
        }

        private string BuildNewMembersBlock(Declaration udt, IEnumerable<(VariableDeclaration Field, string UDTMemberIdentifier)> newMemberPairs)
        {
            var indentation = DetermineIndentationFromLastMember(udt);

            var newMemberStatements = GenerateUserDefinedMemberDeclarations(newMemberPairs, indentation);

            return string.Concat(newMemberStatements);
        }

        private string DetermineIndentationFromLastMember(Declaration udt)
        {
            var lastMember = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.UserDefinedTypeMember)
                .Where(utm => udt == utm.ParentDeclaration)
                .Last();

            lastMember.Context.TryGetPrecedingContext<VBAParser.EndOfStatementContext>(out var endOfStatementContextPrototype);
            return endOfStatementContextPrototype.GetText();
        }

        private IEnumerable<string> GenerateUserDefinedMemberDeclarations(IEnumerable<(VariableDeclaration Field, string UDTMemberIdentifier)> newMemberPairs, string indentation)
            => newMemberPairs.Select(pr => _codeBuilder.UDTMemberDeclaration(pr.UDTMemberIdentifier, pr.Field.AsTypeName, indentation));

        private static void InsertNewMembersBlock(string newMembersBlock, int insertionIndex, IModuleRewriter rewriter) 
            => rewriter.InsertBefore(insertionIndex, $"{newMembersBlock}");

        private int GetInsertionIndex(VBAParser.UdtDeclarationContext udtContext) 
            => udtContext.END_TYPE().Symbol.TokenIndex - 1;
    }
}
