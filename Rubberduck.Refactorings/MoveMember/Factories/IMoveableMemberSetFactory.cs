using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings
{
    public interface IMoveableMemberSetFactory
    {
        IMoveableMemberSet Create(IEnumerable<Declaration> declarations);
    }

    public class MoveableMemberSetFactory : IMoveableMemberSetFactory
    {
        public IMoveableMemberSet Create(IEnumerable<Declaration> declarations)
        {
            return new MoveableMemberSet(declarations);
        }
    }

    public interface IMoveableMemberSetsFactory
    {
        IEnumerable<IMoveableMemberSet> Create(Declaration target);
    }

    public class MoveableMemberSetsFactory : IMoveableMemberSetsFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMoveableMemberSetFactory _factory;
        public MoveableMemberSetsFactory(IDeclarationFinderProvider declarationFinderProvider, IMoveableMemberSetFactory factory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _factory = factory;
        }

        public IEnumerable<IMoveableMemberSet> Create(Declaration target)
        {
            return InitializeMoveableMemberSets(target);
        }

        private IEnumerable<IMoveableMemberSet> InitializeMoveableMemberSets(Declaration moveTarget)
        {
            var groupsByIdentifier = _declarationFinderProvider.DeclarationFinder.Members(moveTarget.QualifiedModuleName)
                    .Where(d => d.IsMember()
                                    || d.IsMemberVariable()
                                    || d.IsModuleConstant()
                                    || d.DeclarationType.Equals(DeclarationType.UserDefinedType)
                                    || d.DeclarationType.Equals(DeclarationType.Enumeration))
                    .GroupBy(key => key.IdentifierName);

            var moveableMembers = new List<IMoveableMemberSet>();
            foreach (var group in groupsByIdentifier)
            {
                var moveableMemberSet = _factory.Create(group.ToList());
                moveableMemberSet.IsSelected = moveableMemberSet.Members.Contains(moveTarget) 
                                        && moveableMemberSet.IdentifierName == moveTarget.IdentifierName;

                var idRefs = new List<IdentifierReference>();
                foreach (var member in moveableMemberSet.Members.Where(m => m.IsMember()))
                {
                    idRefs = FindDirectTypeReferences(member).ToList();

                    var memberContainedReferences = _declarationFinderProvider.DeclarationFinder.IdentifierReferences(member.QualifiedName)
                        .Where(rf => !(rf.Declaration.DeclarationType.HasFlag(DeclarationType.Parameter) || rf.Declaration == rf.ParentScoping));
                    idRefs.AddRange(memberContainedReferences);
                }

                moveableMemberSet.DirectReferences = idRefs;

                moveableMembers.Add(moveableMemberSet);
            }

            var constants = moveableMembers.Where(m => m.Member.IsModuleConstant()).ToList();
            foreach (var moveableMember in constants)
            {
                var lExprContexts = moveableMember.Member.Context.GetDescendents<VBAParser.LExprContext>();
                if (lExprContexts.Any())
                {
                    var otherConstantIdentifierRefs = constants.Where(c => c != moveableMember)
                                                        .SelectMany(oc => oc.Member.References);

                    moveableMember.DirectReferences = otherConstantIdentifierRefs
                                    .Where(rf => lExprContexts.Contains(rf.Context.Parent));
                }
            }

            foreach (var moveableMember in moveableMembers.Where(m => m.Member.IsMemberVariable()).ToList())
            {
                moveableMember.DirectReferences = FindDirectTypeReferences(moveableMember.Member).ToList();
            }
            return moveableMembers;
        }

        private IEnumerable<IdentifierReference> FindDirectTypeReferences(Declaration member)
        {
            var types = _declarationFinderProvider.DeclarationFinder.Members(member.QualifiedModuleName)
                .Where(m => m.DeclarationType.Equals(DeclarationType.UserDefinedType) || m.DeclarationType.Equals(DeclarationType.Enumeration));

            var idRefs = new List<IdentifierReference>();
            foreach (var typeReference in types.AllReferences())
            {
                if (member.AsTypeDeclaration?.Equals(typeReference.Declaration) ?? false)
                {
                    var memberAsTypeContext = member.Context.GetDescendent<VBAParser.AsTypeClauseContext>();
                    var referenceAsTypeContext = typeReference.Context.GetAncestor<VBAParser.AsTypeClauseContext>();
                    if (memberAsTypeContext.Equals(referenceAsTypeContext))
                    {
                        idRefs.Add(typeReference);
                    }
                }
            }
            return idRefs;
        }
    }
}
