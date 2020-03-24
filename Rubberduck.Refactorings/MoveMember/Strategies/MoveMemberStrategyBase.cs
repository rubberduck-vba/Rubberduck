using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.Refactorings.Common;
using System.Diagnostics;
using System;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRefactoringStrategy
    {
        void RefactorRewrite(MoveMemberModel model, IRewriteSession rewriteSession, IRewritingManager rewritingManager, bool asPreview = false);
        IMovedContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, IMovedContentProvider contentToMove);
        bool IsApplicable(MoveMemberModel model);
        bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage);
    }

    public abstract class MoveMemberStrategyBase : IMoveMemberRefactoringStrategy
    {
        public abstract void RefactorRewrite(MoveMemberModel model, IRewriteSession rewriteSession, IRewritingManager rewritingManager, bool asPreview = false);
        public abstract IMovedContentProvider NewDestinationModuleContent(MoveMemberModel model, IRewritingManager rewritingManager, IMovedContentProvider contentToMove);
        public abstract bool IsApplicable(MoveMemberModel model);
        public abstract bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage);

        /// <summary>
        /// Clears entire VariableStmtContext or ConstantStmtContext
        /// when all the variables or constants declared in the list are removed.
        /// </summary>
        /// <param name="rewriteSession"></param>
        /// <param name="declarations"></param>
        protected static void RemoveDeclarations(IRewriteSession rewriteSession, IEnumerable<Declaration> declarations)
        {
            var removedVariables = new Dictionary<VBAParser.VariableListStmtContext, HashSet<Declaration>>();
            var removedConstants = new Dictionary<VBAParser.ConstStmtContext, HashSet<Declaration>>();

            foreach (var declaration in declarations)
            {
                if (declaration.DeclarationType.Equals(DeclarationType.Variable))
                {
                    CacheListDeclaredElement<VBAParser.VariableListStmtContext, VBAParser.VariableSubStmtContext>(rewriteSession, declaration, removedVariables);
                    continue;
                }

                if (declaration.DeclarationType.Equals(DeclarationType.Constant))
                {
                    CacheListDeclaredElement<VBAParser.ConstStmtContext, VBAParser.ConstSubStmtContext>(rewriteSession, declaration, removedConstants);
                    continue;
                }

                var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);
                rewriter.Remove(declaration);
            }

            ExecuteCachedRemoveRequests<VBAParser.VariableListStmtContext, VBAParser.VariableSubStmtContext>(rewriteSession, removedVariables);
            ExecuteCachedRemoveRequests<VBAParser.ConstStmtContext, VBAParser.ConstSubStmtContext>(rewriteSession, removedConstants);
        }

        private static void CacheListDeclaredElement<T, K>(IRewriteSession rewriteSession, Declaration target, Dictionary<T, HashSet<Declaration>> dictionary) where T : ParserRuleContext where K : ParserRuleContext
        {
            var declarationList = target.Context.GetAncestor<T>();

            if ((declarationList?.children.OfType<K>().Count() ?? 1) == 1)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
                rewriter.Remove(target);
                return;
            }

            if (!dictionary.ContainsKey(declarationList))
            {
                dictionary.Add(declarationList, new HashSet<Declaration>());
            }
            dictionary[declarationList].Add(target);
        }

        private static void ExecuteCachedRemoveRequests<T, K>(IRewriteSession rewriteSession, Dictionary<T, HashSet<Declaration>> dictionary) where T : ParserRuleContext where K : ParserRuleContext
        {
            foreach (var key in dictionary.Keys.Where(k => dictionary[k].Any()))
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(dictionary[key].First().QualifiedModuleName);

                if (key.children.OfType<K>().Count() == dictionary[key].Count)
                {
                    rewriter.Remove(key.Parent);
                    continue;
                }

                foreach (var dec in dictionary[key])
                {
                    rewriter.Remove(dec);
                }
            }
        }

        protected static bool DeclarationMoveCreatesNameConflict(Declaration entity, IEnumerable<Declaration> existingModuleDeclarations, IEnumerable<Declaration> declarationMembers = null)
        {
            switch (entity.DeclarationType)
            {
                case DeclarationType.UserDefinedType:
                    return UDTMoveCausesNameConflict(entity, existingModuleDeclarations);
                case DeclarationType.Enumeration:
                    Debug.Assert(declarationMembers != null);
                    return EnumerationMoveCausesNameConflict(entity, existingModuleDeclarations, declarationMembers);
                case DeclarationType.Function:
                case DeclarationType.Procedure:
                case DeclarationType.PropertyGet:
                case DeclarationType.PropertySet:
                case DeclarationType.PropertyLet:
                case DeclarationType.Variable:
                case DeclarationType.Constant:
                    return MemberMoveCausesNameConflict(entity, existingModuleDeclarations);
                default:
                    return true;
            }
        }

        //MS-VBAL 5.3.1.6
        //MS-VBAL 5.3.1.7
        private static bool MemberMoveCausesNameConflict(Declaration member, IEnumerable<Declaration> destinationModuleDeclarations)
        {
            var identifierMatches = destinationModuleDeclarations.Where(pc => member.IdentifierName.IsEquivalentVBAIdentifierTo(pc.IdentifierName));

            if (!identifierMatches.Any()) { return false; }

            if (identifierMatches.Any(idm => idm.IsField() || idm.IsModuleConstant()))
            {
                return true;
            }

            if (member.DeclarationType.HasFlag(DeclarationType.Property) && identifierMatches.Any(idm => idm.DeclarationType.Equals(member.DeclarationType)))
            {
                return true;
            }

            identifierMatches = identifierMatches.Where(idm => !idm.DeclarationType.HasFlag(DeclarationType.Property));
            if (!identifierMatches.Any()) { return false; }

            var memberConflictTypes = new List<DeclarationType>()
            {
                DeclarationType.EnumerationMember,
                DeclarationType.Function,
                DeclarationType.Procedure,
                DeclarationType.LibraryFunction,
                DeclarationType.LibraryProcedure,
            };

            foreach (var declarationTypeFlag in memberConflictTypes)
            {
                if (identifierMatches.Any(idm => idm.DeclarationType.HasFlag(declarationTypeFlag)))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool UDTMoveCausesNameConflict(Declaration udt, IEnumerable<Declaration> destinationModuleDeclarations)
        {
            if (udt.HasPrivateAccessibility())
            {
                return (destinationModuleDeclarations.Any(d => d.IdentifierName.IsEquivalentVBAIdentifierTo(udt.IdentifierName)
                    && (d.DeclarationType.Equals(DeclarationType.UserDefinedType) || d.DeclarationType.Equals(DeclarationType.Enumeration))));
            }
            return false;
        }

        //MS-VBAL 5.2.3.4
        private static bool EnumerationMoveCausesNameConflict(Declaration enumeration, IEnumerable<Declaration> destinationModuleDeclarations, IEnumerable<Declaration> enumMembers)
        {
            var enumerationIdentifierConflictTypes = new List<DeclarationType>()
            {
                DeclarationType.UserDefinedType,
                DeclarationType.Enumeration,
            };

            foreach (var potentialConflict in destinationModuleDeclarations.Where(pc => pc.IdentifierName.IsEquivalentVBAIdentifierTo(enumeration.IdentifierName)))
            {
                if (enumerationIdentifierConflictTypes.Any(ect => ect.HasFlag(potentialConflict.DeclarationType))
                    || potentialConflict.IsField()
                    || potentialConflict.IsModuleConstant())
                {
                    return true;
                }
            }

            var enumMemberIdentifiers = enumMembers.Select(em => em.IdentifierName);
            var identifierMatchingDeclarations 
                            = destinationModuleDeclarations.Where(d => enumMemberIdentifiers
                                    .Contains(d.IdentifierName, StringComparer.InvariantCultureIgnoreCase));

            foreach (var identifierMatch in identifierMatchingDeclarations)
            {
                if (identifierMatch.IsMember()
                    || identifierMatch.IsField()
                    || identifierMatch.IsModuleConstant())
                {
                    return true;
                }
            }
            return false;
        }
    }
}
