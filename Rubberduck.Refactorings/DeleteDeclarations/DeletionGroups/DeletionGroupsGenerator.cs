using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.DeleteDeclarations;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    internal class DeletionGroupsGenerator : IDeclarationDeletionGroupsGenerator
    {
        /// <summary>
        /// DeletionGroupsGeneratorModel represents the data required to generate 1 to n DeclarationDeletionGroups.
        /// Each DeclarationDeletionGroup represents an ordered set of adjacent delete targets within a given scope (module, procedure, etc).
        /// </summary>
        private struct DeletionGroupsGeneratorModel
        {
            public List<IDeclarationDeletionTarget> Targets { set; get; }

            public IOrderedEnumerable<ParserRuleContext> OrderedContexts { set; get; }

            public string AllScopedTargetsDeletedExceptionMessage { set; get; }
        }

        private readonly IDeclarationDeletionGroupFactory _declarationDeletionGroupFactory;

        public DeletionGroupsGenerator(IDeclarationDeletionGroupFactory declarationDeletionGroupFactory)
        {
            _declarationDeletionGroupFactory = declarationDeletionGroupFactory;
        }

        public List<IDeclarationDeletionGroup> Generate(IEnumerable<IDeclarationDeletionTarget> declarationDeletionTargets)
        {
            if (!declarationDeletionTargets.Any())
            {
                return new List<IDeclarationDeletionGroup>();
            }

            var deletionGroupsModels = BuildDeletionGroupsGeneratorModels(declarationDeletionTargets);

            var deletionGroups = new List<IDeclarationDeletionGroup>();
            foreach (var deletionGroupModel in deletionGroupsModels)
            {
                if (!deletionGroupModel.Targets?.Any() ?? true)
                {
                    continue;
                }

                var generatedDeletionGroups = GenerateDeletionGroups(deletionGroupModel, _declarationDeletionGroupFactory);

                deletionGroups.AddRange(generatedDeletionGroups.Where(dg => dg.Declarations.Any()));
            }
            return deletionGroups;
        }

        private static List<DeletionGroupsGeneratorModel> BuildDeletionGroupsGeneratorModels(IEnumerable<IDeclarationDeletionTarget> deleteDeclarationTargets)
        {
            var deletionGroupsModels = new List<DeletionGroupsGeneratorModel>();

            if (ContainsModuleScopeTargets(deleteDeclarationTargets, out var moduleScopeTargets))
            {
                var moduleScopeModel = new DeletionGroupsGeneratorModel()
                {
                    Targets = moduleScopeTargets,
                    OrderedContexts = GetOrderedContextsForModuleScopedTargets(moduleScopeTargets.First()),
                };

                deletionGroupsModels.Add(moduleScopeModel);

                deleteDeclarationTargets = deleteDeclarationTargets.Except(moduleScopeModel.Targets);
            }

            if (ContainsLocalScopeTargets(deleteDeclarationTargets, out var procedureScopeTargets))
            {
                var localScopeModels = CreateNonModuleScopedDeletionGroupModels(
                    procedureScopeTargets,
                    (t) => t.TargetProxy.ParentDeclaration.Context);

                deletionGroupsModels.AddRange(localScopeModels);

                deleteDeclarationTargets = deleteDeclarationTargets.Except(procedureScopeTargets);
            }

            if (ContainsUnterminatedBlockScopeTargets(deleteDeclarationTargets, out var unterminatedBlockScopeTargets))
            {
                var unterminatedBlockScopeModels = CreateNonModuleScopedDeletionGroupModels(
                    unterminatedBlockScopeTargets,
                    (t) => t.TargetContext.Parent as ParserRuleContext);

                deletionGroupsModels.AddRange(unterminatedBlockScopeModels);

                deleteDeclarationTargets = deleteDeclarationTargets.Except(unterminatedBlockScopeTargets);
            }

            if (ContainsEnumerationMembers(deleteDeclarationTargets, out var enumMembers))
            {
                var enumerationMembersModels = CreateNonModuleScopedDeletionGroupModels(
                    enumMembers,
                    (t) => t.TargetProxy.ParentDeclaration.Context,
                    "At least one EnumerationMember must be retained");

                deletionGroupsModels.AddRange(enumerationMembersModels);

                deleteDeclarationTargets = deleteDeclarationTargets.Except(enumMembers);
            }

            if (ContainsUDTMembers(deleteDeclarationTargets, out var udtMembers))
            {
                var userDefinedTypeMemberModels = CreateNonModuleScopedDeletionGroupModels(
                    udtMembers,
                    (t) => t.TargetProxy.ParentDeclaration.Context,
                    "At least one UserDefinedTypeMember must be retained");

                deletionGroupsModels.AddRange(userDefinedTypeMemberModels);

                deleteDeclarationTargets = deleteDeclarationTargets.Except(udtMembers);
            }

            if (deleteDeclarationTargets.Any())
            {
                throw new InvalidOperationException($"Encountered Unhandled Target:{deleteDeclarationTargets.First().TargetProxy.DeclarationType}({deleteDeclarationTargets.First().TargetProxy.IdentifierName})");
            }

            return deletionGroupsModels;
        }

        private static List<IDeclarationDeletionGroup> GenerateDeletionGroups(DeletionGroupsGeneratorModel model, IDeclarationDeletionGroupFactory declarationDeletionGroupFactory)
        {
            var orderedContexts = model.OrderedContexts.ToList();

            var nonDeleteIndices = orderedContexts.Where(c => !model.Targets.Any(d => d.TargetContext == c))
                .Select(c => orderedContexts.IndexOf(c)).OrderBy(t => t).ToList();

            //If there are zero nonDeleteIndices, the request is to Delete all the declarations in 
            //the scope -> return a single DeletionGroup
            if (!nonDeleteIndices.Any())
            {
                if (model.AllScopedTargetsDeletedExceptionMessage != null)
                {
                    //Deleting all EnumerationMembers of an Enumeration or all UserDefinedTypeMembers of a UserDefinedType
                    //results in uncompilable code.  It is an exception here because this use case should have been handled upstream.
                    throw new InvalidOperationException(model.AllScopedTargetsDeletedExceptionMessage);
                }

                return new List<IDeclarationDeletionGroup>()
                {
                    declarationDeletionGroupFactory.Create(model.Targets.OrderBy(t => t.TargetContext.GetSelection()))
                };
            }

            var deletionGroups = new List<IDeclarationDeletionGroup>();

            foreach ((int? Start, int? End) indexPair in NonDeleteIndicePairGenerator.Generate(nonDeleteIndices))
            {
                if (!(indexPair.Start.HasValue || indexPair.End.HasValue))
                {
                    continue;
                }

                var groupedTargets = GetOrderedDeleteTargets(model, (indexPair.Start, indexPair.End));

                var deletionGroup = declarationDeletionGroupFactory.Create(groupedTargets);

                deletionGroup.PrecedingNonDeletedContext = indexPair.Start.HasValue
                    ? orderedContexts.ElementAt(indexPair.Start.Value)
                    : null;

                deletionGroups.Add(deletionGroup);
            }

            return deletionGroups;
        }

        private static IOrderedEnumerable<IDeclarationDeletionTarget> GetOrderedDeleteTargets(DeletionGroupsGeneratorModel model, (int? Start, int? End) nonDeletePair)
        {
            var deleteContexts = new List<ParserRuleContext>();

            //DeletionTargets occupy the first 1 to n contexts preceding the first retained context in scope
            if (!nonDeletePair.Start.HasValue)
            {
                deleteContexts = model.OrderedContexts.Take(nonDeletePair.End.Value).ToList();
            }
            //DeletionTargets occupy the last 1 to n contexts after the last retained context in scope
            else if (!nonDeletePair.End.HasValue)
            {
                deleteContexts = model.OrderedContexts.Skip(nonDeletePair.Start.Value + 1).ToList();
            }
            //A group of delete target contexts are bounded by a pair of retained contexts in scope
            else
            {
                deleteContexts = model.OrderedContexts
                    .Skip(nonDeletePair.Start.Value + 1)
                    .Take(nonDeletePair.End.Value - nonDeletePair.Start.Value - 1)
                    .ToList();
            }

            return model.Targets.Where(t => deleteContexts
                .Contains(t.TargetContext))
                .OrderBy(rt => rt.TargetContext.GetSelection());
        }

        private static IEnumerable<DeletionGroupsGeneratorModel> CreateNonModuleScopedDeletionGroupModels(IEnumerable<IDeclarationDeletionTarget> targets, Func<IDeclarationDeletionTarget, ParserRuleContext> getOrganizingContext, string allScopedTargetsDeletedExceptionMessage = null)
        {
            var deletionGroupModels = new List<DeletionGroupsGeneratorModel>();

            foreach (var targetGroup in targets.ToLookup(t => getOrganizingContext(t)))
            {
                var deletionGroupModel = new DeletionGroupsGeneratorModel()
                {
                    Targets = targetGroup.ToList(),
                    OrderedContexts = GetOrderedContextsForNonModuleScopedTargets(targetGroup.First()),
                    AllScopedTargetsDeletedExceptionMessage = allScopedTargetsDeletedExceptionMessage
                };

                deletionGroupModels.Add(deletionGroupModel);
            }

            return deletionGroupModels;
        }

        private static bool ContainsModuleScopeTargets(IEnumerable<IDeclarationDeletionTarget> deleteDeclarationTargets, out List<IDeclarationDeletionTarget> moduleScopeTargets)
        {
            moduleScopeTargets = deleteDeclarationTargets.Where(t => t.TargetProxy.ParentDeclaration is ModuleDeclaration).ToList();
            return moduleScopeTargets.Any();
        }

        private static bool ContainsLocalScopeTargets(IEnumerable<IDeclarationDeletionTarget> deleteDeclarationTargets, out List<IDeclarationDeletionTarget> procedureScopeTargets)
        {
            procedureScopeTargets = deleteDeclarationTargets.Where(t => t.TargetProxy.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member)
                && !(t.TargetContext.Parent is VBAParser.UnterminatedBlockContext)).ToList();

            return procedureScopeTargets.Any();
        }

        private static bool ContainsUnterminatedBlockScopeTargets(IEnumerable<IDeclarationDeletionTarget> deleteDeclarationTargets, out List<IDeclarationDeletionTarget> unterminatedBlockScopeTargets)
        {
            unterminatedBlockScopeTargets = deleteDeclarationTargets.Where(t => t.TargetProxy.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Member)
                && t.TargetContext.Parent is VBAParser.UnterminatedBlockContext).ToList();

            return unterminatedBlockScopeTargets.Any();
        }

        private static bool ContainsUDTMembers(IEnumerable<IDeclarationDeletionTarget> deleteDeclarationTargets, out List<IDeclarationDeletionTarget> udtMembers)
        {
            udtMembers = deleteDeclarationTargets.Where(t => t.TargetProxy.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember)).ToList();
            return udtMembers.Any();
        }

        private static bool ContainsEnumerationMembers(IEnumerable<IDeclarationDeletionTarget> deleteDeclarationTargets, out List<IDeclarationDeletionTarget> enumMembers)
        {
            enumMembers = deleteDeclarationTargets.Where(t => t.TargetProxy.DeclarationType.HasFlag(DeclarationType.EnumerationMember)).ToList();
            return enumMembers.Any();
        }

        private static IOrderedEnumerable<ParserRuleContext> GetOrderedContextsForNonModuleScopedTargets(IDeclarationDeletionTarget deleteTarget)
        {
            if (deleteTarget is ILocalScopeDeletionTarget localTarget)
            {
                return localTarget.ScopingContext.GetChildrenOfType<VBAParser.BlockStmtContext>()
                    .OrderBy(e => e?.GetSelection());
            }

            var contexts = deleteTarget.TargetProxy.DeclarationType == DeclarationType.EnumerationMember
                ? deleteTarget.TargetProxy.Context.GetAncestor<VBAParser.EnumerationStmtContext>().GetChildrenOfType<VBAParser.EnumerationStmt_ConstantContext>()
                : deleteTarget.TargetProxy.Context.GetAncestor<VBAParser.UdtMemberListContext>().GetChildrenOfType<VBAParser.UdtMemberContext>();
            
            return contexts.OrderBy(e => e?.GetSelection());
        }

        private static IOrderedEnumerable<ParserRuleContext> GetOrderedContextsForModuleScopedTargets(IDeclarationDeletionTarget deleteTarget)
        {
            var moduleContext = deleteTarget.TargetProxy.Context.GetAncestor<VBAParser.ModuleContext>();

            var moduleDeclarationElements = moduleContext.GetChild<VBAParser.ModuleDeclarationsContext>()
                .GetChildrenOfType<VBAParser.ModuleDeclarationsElementContext>();

            var moduleBodyElements = moduleContext.GetChild<VBAParser.ModuleBodyContext>()
                .GetChildrenOfType<VBAParser.ModuleBodyElementContext>();

            return moduleDeclarationElements.Concat(moduleBodyElements)
                .OrderBy(c => c?.GetSelection());
        }
    }
}
