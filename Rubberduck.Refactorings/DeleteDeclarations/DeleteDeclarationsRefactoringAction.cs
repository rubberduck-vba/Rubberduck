using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    /// <summary>
    /// Removes 0 to n Declarations along with associated Annotations and comments on the same logicalline
    /// of the Declaration.  Other surrounding comments are edited to include a TODO statement indicating
    /// that the user should evaluate if the comment is still valid.
    /// Warning: If references to the removed Declarations are not removed/modified by other code, 
    /// this refactoring action will generate uncompilable code.
    /// </summary>
    public class DeleteDeclarationsRefactoringAction : CodeOnlyRefactoringActionBase<DeleteDeclarationsModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IRewritingManager _rewritingManager;
        private static List<DeclarationType> _supportedDeclarationTypes = new List<DeclarationType>()
        {
            DeclarationType.Variable,
            DeclarationType.Constant,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet,
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember,
            DeclarationType.Enumeration,
            DeclarationType.EnumerationMember,
            DeclarationType.LineLabel
        };

        public DeleteDeclarationsRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
            : base(rewritingManager) 
        {
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
        }

        public override void Refactor(DeleteDeclarationsModel model, IRewriteSession rewriteSession)
        {
            if (!model.Targets.Any())
            {
                return;
            }

            if (model.Targets.Any(t => !_supportedDeclarationTypes.Contains(t.DeclarationType)))
            {
                var invalidDeclaration = model.Targets.First(t => !_supportedDeclarationTypes.Contains(t.DeclarationType));
                throw new InvalidDeclarationTypeException(invalidDeclaration);
            }

            var scrubbedTargets = ModifyTargetsListToAvoidBoundaryReplacementErrors(model);

            DeleteModuleElements(scrubbedTargets, rewriteSession);
            DeleteProcedureScopeElements(scrubbedTargets, rewriteSession);
            DeleteUserDefinedTypeMembers(scrubbedTargets, rewriteSession);
            DeleteEnumerationMembers(scrubbedTargets, rewriteSession);
        }

        private void DeleteModuleElements(IEnumerable<Declaration> targets, IRewriteSession rewriteSession)
        {
            var moduleElementTargets = targets.Where(t => t.ParentDeclaration is ModuleDeclaration).ToList();
            if (moduleElementTargets.Any())
            {
                var refactoringAction = new DeleteModuleElementsRefactoringAction(_declarationFinderProvider, _rewritingManager);
                refactoringAction.Refactor(new DeleteModuleElementsModel(moduleElementTargets), rewriteSession);
            }
        }

        private void DeleteProcedureScopeElements(IEnumerable<Declaration> targets, IRewriteSession rewriteSession)
        {
            var procedureLocalTargets = targets.Where(t => !(t.ParentDeclaration is ModuleDeclaration)
                && !(t.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember) || t.DeclarationType.HasFlag(DeclarationType.EnumerationMember)))
                .ToList();

            if (procedureLocalTargets.Any())
            {
                var refactoringAction = new DeleteProcedureScopeElementsRefactoringAction(_declarationFinderProvider, _rewritingManager);
                refactoringAction.Refactor(new DeleteProcedureScopeElementsModel(procedureLocalTargets), rewriteSession);
            }
        }
        private void DeleteUserDefinedTypeMembers(IEnumerable<Declaration> targets, IRewriteSession rewriteSession)
        {
            var udtMemberTargets = targets.Where(t => t.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember)).ToList();
            if (udtMemberTargets.Any())
            {
                var refactoringAction = new DeleteUDTMembersRefactoringAction(_declarationFinderProvider, _rewritingManager);
                refactoringAction.Refactor(new DeleteUDTMembersModel(udtMemberTargets), rewriteSession);
            }
        }

        private void DeleteEnumerationMembers(IEnumerable<Declaration> targets, IRewriteSession rewriteSession)
        {
            var enumMemberTargets = targets.Where(t => t.DeclarationType.HasFlag(DeclarationType.EnumerationMember)).ToList();
            if (enumMemberTargets.Any())
            {
                var refactoringAction = new DeleteEnumMembersRefactoringAction(_declarationFinderProvider, _rewritingManager);
                refactoringAction.Refactor(new DeleteEnumMembersModel(enumMemberTargets), rewriteSession);
            }
        }

        private static List<Declaration> ModifyTargetsListToAvoidBoundaryReplacementErrors(DeleteDeclarationsModel model)
        {
            var targets = model.Targets.ToList();

            targets = RemoveTargetChildren(targets);

            List<Declaration> toRemove = new List<Declaration>();
            List<Declaration> toAdd = new List<Declaration>();

            //Replace members with the Parent declaration if all the members are in the list of targets
            var modifiesEnumMemberTargets = RequiresEnumDeclarationDeletion(targets, ref toRemove, ref toAdd);
            var modifiesUDTMemberTargets = RequiresUserDefinedTypeDeclarationDeletion(targets, ref toRemove, ref toAdd);

            if (modifiesEnumMemberTargets || modifiesUDTMemberTargets)
            {
                targets.RemoveAll(t => toRemove.Contains(t));
                targets.AddRange(toAdd);
            }

            return targets;
        }

        //Remove targets where the Parent declaration is also in the list of deletion targets
        private static List<Declaration> RemoveTargetChildren(List<Declaration> targets)
        {
            var declarationTypes = new List<DeclarationType>() 
            { 
                DeclarationType.Member,
                DeclarationType.Enumeration,
                DeclarationType.UserDefinedType
            };

            foreach (var decType in declarationTypes)
            {
                var parentDeclarations = targets.Where(t => t.DeclarationType.HasFlag(decType));
                var toRemove = targets.Where(t => parentDeclarations.Contains(t.ParentDeclaration));
                targets.RemoveAll(t => toRemove.Contains(t));
            }
            return targets;
        }

        private static bool RequiresEnumDeclarationDeletion(List<Declaration> targets, ref List<Declaration> toRemove, ref List<Declaration> toAdd)
        {

            bool ContainsAllMembers(ParserRuleContext ctxt, IEnumerable<Declaration> declarations)
                => ctxt.children.Where(ch => ch is VBAParser.EnumerationStmt_ConstantContext).Count() == declarations.Count();

            var enumMembers = targets.Where(t => t.DeclarationType == DeclarationType.EnumerationMember);
            return RequiresParentDeclarationDeletion(enumMembers, ContainsAllMembers, ref toRemove, ref toAdd);
        }

        private static bool RequiresUserDefinedTypeDeclarationDeletion(List<Declaration> targets, ref List<Declaration> toRemove, ref List<Declaration> toAdd)
        {

            bool ContainsAllMembers(ParserRuleContext ctxt, IEnumerable<Declaration> declarations)
                => ctxt.GetChild<VBAParser.UdtMemberListContext>()
                    .children.Where(ch => ch is VBAParser.UdtMemberContext).Count() == declarations.Count();

            var udtMembers = targets.Where(t => t.DeclarationType == DeclarationType.UserDefinedTypeMember);
            return RequiresParentDeclarationDeletion(udtMembers, ContainsAllMembers, ref toRemove, ref toAdd);
        }

        private static bool RequiresParentDeclarationDeletion(IEnumerable<Declaration> targets, Func<ParserRuleContext, IEnumerable<Declaration>, bool> requiresParentDeletion, ref List<Declaration> toRemove, ref List<Declaration> toAdd)
        {
            foreach (var tGroup in targets.ToLookup(key => key.ParentDeclaration))
            {
                if (requiresParentDeletion(tGroup.Key.Context, tGroup))
                {
                    toRemove.AddRange(tGroup);
                    toAdd.Add(tGroup.Key);
                }
            }

            return toRemove.Count > 0;
        }
    }
}
