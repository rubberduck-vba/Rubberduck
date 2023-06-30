using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    /// <summary>
    /// Removes 0 to n Declarations along with associated Annotations. Removes comments on the same logical line
    /// as the removed Declaration.  Other surrounding comments are edited to include a TODO statement indicating
    /// that the user should evaluate if the comment is still valid.
    /// Warning: If references to the removed Declarations are not removed/modified by other code, 
    /// this refactoring action will generate uncompilable code.
    /// </summary>
    public class DeleteDeclarationsRefactoringAction : CodeOnlyRefactoringActionBase<DeleteDeclarationsModel>
    {
        private readonly ICodeOnlyRefactoringAction<DeleteModuleElementsModel> _deleteModuleElementsRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<DeleteProcedureScopeElementsModel> _deleteProcedureScopeElementsRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<DeleteUDTMembersModel> _deleteUDTMembersRefactoringAction;
        private readonly ICodeOnlyRefactoringAction<DeleteEnumMembersModel> _deleteEnumMembersRefactoringAction;

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

        public DeleteDeclarationsRefactoringAction(DeleteModuleElementsRefactoringAction deleteModuleElementsRefactoringAction,
            DeleteProcedureScopeElementsRefactoringAction deleteProcedureScopeElementsRefactoringAction,
            DeleteUDTMembersRefactoringAction deleteUDTMembersRefactoringAction,
            DeleteEnumMembersRefactoringAction deleteEnumMembersRefactoringAction,
            IRewritingManager rewritingManager)
            : base(rewritingManager)
        {
            _deleteModuleElementsRefactoringAction = deleteModuleElementsRefactoringAction;
            _deleteProcedureScopeElementsRefactoringAction = deleteProcedureScopeElementsRefactoringAction;
            _deleteUDTMembersRefactoringAction = deleteUDTMembersRefactoringAction;
            _deleteEnumMembersRefactoringAction = deleteEnumMembersRefactoringAction;
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

            //Minimize/optimize target list to prevent Rewriter boundary overlap errors and uncompilable code scenarios
            var minimizedTargets = RemoveTargetChildren(model.Targets);
            minimizedTargets = ReplaceDeleteAllUDTMembersOrEnumMembersWithParent(minimizedTargets);

            var targetsGroup = GroupTargetsByRefactoringActionScope(minimizedTargets);

            DeleteTargets<DeleteModuleElementsModel>(targetsGroup.ModuleScope, model, rewriteSession);
            DeleteTargets<DeleteProcedureScopeElementsModel>(targetsGroup.ProcedureScope, model, rewriteSession);
            DeleteTargets<DeleteEnumMembersModel>(targetsGroup.EnumMembers, model, rewriteSession);
            DeleteTargets<DeleteUDTMembersModel>(targetsGroup.UdtMembers, model, rewriteSession);
        }

        private void DeleteTargets<T>(IEnumerable<Declaration> targets, DeleteDeclarationsModel model, IRewriteSession rewriteSession) where T : DeleteDeclarationsModel, new()
        {
            if (!targets.Any())
            {
                return;
            }

            var tModel = CreateCodeOnlyRefactoringModel<T>(targets, model);

            switch (tModel)
            {
                case DeleteModuleElementsModel deleteModuleElementsModel:
                    _deleteModuleElementsRefactoringAction.Refactor(deleteModuleElementsModel, rewriteSession);
                    return;
                case DeleteProcedureScopeElementsModel procedureScopeElementsModel:
                    _deleteProcedureScopeElementsRefactoringAction.Refactor(procedureScopeElementsModel, rewriteSession);
                    return;
                case DeleteEnumMembersModel enumMembersModel:
                    _deleteEnumMembersRefactoringAction.Refactor(enumMembersModel, rewriteSession);
                    return;
                case DeleteUDTMembersModel udtMembersModel:
                    _deleteUDTMembersRefactoringAction.Refactor(udtMembersModel, rewriteSession);
                    return;
                default:
                    throw new ArgumentException();
            }
        }

        private static (IEnumerable<Declaration> ModuleScope,
            IEnumerable<Declaration> ProcedureScope,
            IEnumerable<Declaration> EnumMembers,
            IEnumerable<Declaration> UdtMembers)
        GroupTargetsByRefactoringActionScope(IEnumerable<Declaration> targets)
        {
            var moduleScopeTargets = targets.Where(t => t.ParentDeclaration is ModuleDeclaration);

            var enumMemberTargets = targets.Where(t => t.DeclarationType.HasFlag(DeclarationType.EnumerationMember));

            var udtMemberTargets = targets.Where(t => t.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember));

            var procedureScopeTargets = targets
                .Except(moduleScopeTargets)
                .Except(enumMemberTargets)
                .Except(udtMemberTargets);

            return (moduleScopeTargets, procedureScopeTargets, enumMemberTargets, udtMemberTargets);
        }

        /// <summary>
        /// Removes targets where the Parent declaration is also in the list of deletion targets
        /// </summary>
        private static List<Declaration> RemoveTargetChildren(IEnumerable<Declaration> targets)
        {
            var declarationTypes = new List<DeclarationType>() 
            { 
                DeclarationType.Member,
                DeclarationType.Enumeration,
                DeclarationType.UserDefinedType
            };

            var optimizedTargets = targets.ToList();

            foreach (var decType in declarationTypes)
            {
                var parentDeclarations = targets.Where(t => t.DeclarationType.HasFlag(decType));
                var toRemove = targets.Where(t => parentDeclarations.Contains(t.ParentDeclaration));
                optimizedTargets.RemoveAll(t => toRemove.Contains(t));
            }
            return optimizedTargets;
        }

        /// <summary>
        /// If all members of an Enum or UserDefinedType are targeted for deletion, adds the UserDefinedType
        /// and/or Enum declaration to the target list and removes the associated Members.  Both UDT and
        /// Enum Types must have at least one member to compile.
        /// </summary>
        private static List<Declaration> ReplaceDeleteAllUDTMembersOrEnumMembersWithParent(List<Declaration> targets)
        {
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

        private static T CreateCodeOnlyRefactoringModel<T>(IEnumerable<Declaration> targets, DeleteDeclarationsModel model) where T: DeleteDeclarationsModel, new()
        {
            var newModel = new T()
            {
                InsertValidationTODOForRetainedComments = model.InsertValidationTODOForRetainedComments,
                DeleteDeclarationLogicalLineComments = model.DeleteDeclarationLogicalLineComments,
                DeleteAnnotations = model.DeleteAnnotations,
                DeleteDeclarationsOnly = model.DeleteDeclarationsOnly
            };

            newModel.AddRangeOfDeclarationsToDelete(targets);
            return newModel;
        }
    }
}
