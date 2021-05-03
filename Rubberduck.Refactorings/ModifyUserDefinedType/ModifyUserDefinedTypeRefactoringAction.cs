using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.ModifyUserDefinedType
{
    public class ModifyUserDefinedTypeRefactoringAction : CodeOnlyRefactoringActionBase<ModifyUserDefinedTypeModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IRewritingManager _rewritingManager;
        private readonly ICodeBuilder _codeBuilder;

        /// <summary>
        /// Removes or adds UserDefinedTypeMember declarations to an existing UserDefinedType. 
        /// Adding a UDTMember is based on a Declaration prototype (typically a variable declaration but can a UserDefinedTypeMember, Constant, or Function).
        /// </summary>
        /// <remarks>
        /// The refactoring actions does not modify the prototype declaration or its references.
        /// The refactoring actions does not modify references for removed UDTMembers.
        /// The refactoring action does not provide any identifier validation or conflictAnalysis
        /// </remarks>
        public ModifyUserDefinedTypeRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager, ICodeBuilder codeBuilder)
            :base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
            _codeBuilder = codeBuilder;
        }

        public override void Refactor(ModifyUserDefinedTypeModel model, IRewriteSession rewriteSession)
        {
            var newMembers = new List<string>();
            foreach ((Declaration Prototype, string Identifier) in model.MembersToAdd)
            {
                _codeBuilder.TryBuildUDTMemberDeclaration(Prototype, Identifier, out var udtMemberDeclaration);
                newMembers.Add(udtMemberDeclaration);
            }

            var scratchPad = _rewritingManager.CheckOutCodePaneSession().CheckOutModuleRewriter(model.Target.QualifiedModuleName);
            scratchPad.InsertBefore(model.InsertionIndex, $"{Environment.NewLine}{string.Join(Environment.NewLine, newMembers)}");

            foreach (var member in model.MembersToRemove)
            {
                scratchPad.Remove(member);
            }

            var udtDeclarationContext = model.Target.Context as VBAParser.UdtDeclarationContext;
            var newBlock = scratchPad.GetText(udtDeclarationContext.Start.TokenIndex, udtDeclarationContext.Stop.TokenIndex);
            var udtLines = newBlock.Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
                .Where(ul => !string.IsNullOrEmpty(ul.Trim()));

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.Target.QualifiedModuleName);
            rewriter.Replace(udtDeclarationContext, string.Join(Environment.NewLine, _codeBuilder.Indenter.Indent(udtLines)));
        }
    }
}
