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
    /// <summary>
    /// CreateUDTMemberRefactoringAction adds a UserDefinedTypeMember declaration (based on a
    /// prototype declaation) to a UserDefinedType declaration.  The indentation of the first 
    /// existing member is used by the inserted members.  The caller is responsible for identifier validation and name collision anaysis.
    /// </summary>
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

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.UserDefinedTypeTargets.First().QualifiedModuleName);

            foreach (var udt in model.UserDefinedTypeTargets)
            {
                var newMembersBlock = BuildNewMembersBlock(udt, model);

                var insertionIndex = (udt.Context as VBAParser.UdtDeclarationContext)
                    .END_TYPE().Symbol.TokenIndex - 1;

                rewriter.InsertBefore(insertionIndex, $"{newMembersBlock}");
            }
        }

        private string BuildNewMembersBlock(Declaration udt, CreateUDTMemberModel model)
        {
            var indentation = DetermineIndentationFromLastMember(udt);

            var newMemberStatements = GenerateUserDefinedMemberDeclarations(model[udt], indentation);

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

        private IEnumerable<string> GenerateUserDefinedMemberDeclarations(IEnumerable<(Declaration Prototype, string UDTMemberIdentifier)> newMemberPairs, string indentation)
        {
            var declarations = new List<string>();
            foreach (var (Prototype, UDTMemberIdentifier) in newMemberPairs)
            {
                if (_codeBuilder.TryBuildUDTMemberDeclaration(UDTMemberIdentifier, Prototype, out var declaration))
                {
                    declarations.Add($"{indentation}{declaration}");
                }
            }
            return declarations;
        }
    }
}
