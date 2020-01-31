using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberContentInfo
    {
        private IProvideMoveDeclarationGroups _moveSource;
        private Dictionary<Declaration, MemberContentInfo> _memberContent;

        public MoveMemberContentInfo(IProvideMoveDeclarationGroups moveSource)
        {
            _moveSource = moveSource;
            _memberContent = new Dictionary<Declaration, MemberContentInfo>();
        }

        public IEnumerable<IdentifierReference> ReferenceCandidiatesToPassAsArguments(Declaration member)
        {
            var memberContentInfo = RetrieveContentInfo(member);
            return ReferenceCandidiatesToPassAsArguments(memberContentInfo);
        }

        private IEnumerable<IdentifierReference> ReferenceCandidiatesToPassAsArguments(MemberContentInfo memberContentInfo)
        {
            var argumentIdentifierReferences = new List<IdentifierReference>();

            var idxExprContextsFilter = new List<VBAParser.IndexExprContext>();
            var firstPassRefs = memberContentInfo.ModuleScopeNonLocalReferences.Where(rf => !rf.Declaration.DeclarationType.Equals(DeclarationType.Procedure))
                .Where(rf => !_moveSource.MoveAndDelete.Contains(rf.Declaration));

            foreach (var memberIdRefs in firstPassRefs.Where(rf => rf.Declaration.IsMember()))
            {
                //TODO: Test for nested function calls...some member references would need to be removed too...yes?
                //e.g. Function DoSomething(Func1(arg1, PropertyGet1, Func2(PropertGet2)), arg3, ar4) As Long
                //becomes at call site (Foo arg1, ag2 uses DoSomething within body): =>Sub Foo(arg1, arg2), becomes Sub Foo(ByVal arg_DoSomething As Long, arg1, arg2)
                argumentIdentifierReferences.Add(memberIdRefs);
                if (memberIdRefs.Context.TryGetAncestor<VBAParser.IndexExprContext>(out var idxExpr))
                {
                    idxExprContextsFilter.Add(idxExpr);
                }
            }

            foreach (var nonMemberIdRef in firstPassRefs.Where(rf => !rf.Declaration.IsMember()))
            {
                if (nonMemberIdRef.Context.TryGetAncestor<VBAParser.IndexExprContext>(out var idxExpr))
                {
                    if (!idxExprContextsFilter.Contains(idxExpr))
                    {
                        argumentIdentifierReferences.Add(nonMemberIdRef);
                    }
                }
                else
                {
                    argumentIdentifierReferences.Add(nonMemberIdRef);
                }
            }
            return argumentIdentifierReferences;
        }

        public string CallSiteArguments(Declaration member)
        {
            var memberContentInfo = RetrieveContentInfo(member);
            return BuildParameterList(member, memberContentInfo.ForwardingCallSiteArgs, (VBAParser.ArgContext arg) => arg.GetChild<VBAParser.UnrestrictedIdentifierContext>().GetText());
        }

        public string DestinationSignatureParameters(Declaration member)
        {
            var memberContentInfo = RetrieveContentInfo(member);
            return BuildParameterList(member, memberContentInfo.DestinationSignatureArgs, (VBAParser.ArgContext arg) => arg.GetText());
        }

        private MemberContentInfo RetrieveContentInfo(Declaration member)
        {
            if (!_memberContent.ContainsKey(member))
            {
                var memberContentInfo = new MemberContentInfo(member, _moveSource.AllDeclarations);
                memberContentInfo = InitializeArgumentCollections(memberContentInfo, _moveSource.MoveAndDelete.AllDeclarations);
                _memberContent.Add(member, memberContentInfo);
            }
            return _memberContent[member];
        }

        private MemberContentInfo InitializeArgumentCollections(MemberContentInfo memberContentInfo, IEnumerable<Declaration> movingDeclarations)
        {
            var forwardingCallSiteArgs_ByRef = new List<string>();
            var forwardingCallSiteArgs_ByVal = new List<string>();
            var destinationSignatureArgs_ByVal = new List<string>();
            var destinationSignatureArgs_ByRef = new List<string>();

            var identifierRefsOfInterest = memberContentInfo.ModuleScopeNonLocalReferences //.ModuleScopeNonLocalWriteReferences.Concat(memberContentInfo.ModuleScopeNonLocalReadReferences)
                .Where(rf => !movingDeclarations.Contains(rf.Declaration)
                    && !(rf.Declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.Module) ?? false));

            var writeIDRefsOfInterest = identifierRefsOfInterest.Where(rf => rf.IsAssignment);
            var readIDRefsOfInterest = identifierRefsOfInterest.Except(writeIDRefsOfInterest).Where(rf => !rf.Declaration.DeclarationType.Equals(DeclarationType.Procedure));

            var IDRefComparer = new IdentifierReferenceComparer();
            foreach (var assignment in writeIDRefsOfInterest.Distinct(IDRefComparer).OrderBy(rf => rf.IdentifierName))
            {
                forwardingCallSiteArgs_ByRef.Add(assignment.IdentifierName);
                destinationSignatureArgs_ByRef.Add(ArgSignature(Tokens.ByRef, assignment));
            }

            foreach (var read in readIDRefsOfInterest.Distinct(IDRefComparer).OrderBy(rf => rf.IdentifierName))
            {
                //If read AND write references exist, pass only ByRef
                if (forwardingCallSiteArgs_ByRef.Contains(read.IdentifierName))
                {
                    continue;
                }

                var callSiteArg = read.IdentifierName;
                if (read.Declaration.DeclarationType.HasFlag(DeclarationType.Function))
                {
                    callSiteArg = read.Context.Parent.GetText();
                }
                forwardingCallSiteArgs_ByVal.Add(callSiteArg);
                destinationSignatureArgs_ByVal.Add(ArgSignature(Tokens.ByVal, read));
            }

            memberContentInfo.ForwardingCallSiteArgs = forwardingCallSiteArgs_ByVal.Concat(forwardingCallSiteArgs_ByRef).ToList();
            memberContentInfo.DestinationSignatureArgs = destinationSignatureArgs_ByVal.Concat(destinationSignatureArgs_ByRef).ToList();
            return memberContentInfo;
        }

        private string ArgSignature(string ByRefByVal, IdentifierReference idRef)
            => $"{ByRefByVal} {MoveMemberResources.Prefix_Parameter}{idRef.IdentifierName} {Tokens.As} {idRef.Declaration.AsTypeName}";

        private class IdentifierReferenceComparer : IEqualityComparer<IdentifierReference>
        {
            public bool Equals(IdentifierReference lhs, IdentifierReference rhs)
            {
                return lhs.IdentifierName.Equals(rhs.IdentifierName);
            }

            public int GetHashCode(IdentifierReference obj)
            {
                return obj.IdentifierName.GetHashCode();
            }
        }

        private string BuildParameterList(Declaration d, IEnumerable<string> insertAtFront, Func<VBAParser.ArgContext, string> argContentExtraction)
        {
            var argContexts = d.Context.GetDescendents<VBAParser.ArgContext>();
            var allArgs = insertAtFront.Concat(argContexts.Select(arg => argContentExtraction(arg)));
            return string.Join(", ", allArgs);
        }

        private struct MemberContentInfo
        {
            private readonly List<Declaration> _moduleDeclarations;

            public MemberContentInfo(Declaration member, IEnumerable<Declaration> moduleDeclarations)
            {
                _moduleDeclarations = moduleDeclarations.ToList();
                _moduleScopeNonLocalDeclarationRefs = null;
                Member = member;
                ForwardingCallSiteArgs = new List<string>();
                DestinationSignatureArgs = new List<string>();
            }

            public Declaration Member { get; }

            public List<string> ForwardingCallSiteArgs { set; get; }
            public List<string> DestinationSignatureArgs { set; get; }

            private List<IdentifierReference> ModuleScopeNonLocalReadReferences
                => ModuleScopeNonLocalDeclarationReferencesWithinMember.Where(rf => !rf.IsAssignment).ToList();

            private List<IdentifierReference> ModuleScopeNonLocalWriteReferences
                => ModuleScopeNonLocalDeclarationReferencesWithinMember.Where(rf => rf.IsAssignment).ToList();

            public List<IdentifierReference> ModuleScopeNonLocalReferences
                => ModuleScopeNonLocalWriteReferences.Concat(ModuleScopeNonLocalReadReferences).ToList();

            private List<IdentifierReference> _moduleScopeNonLocalDeclarationRefs;
            public IEnumerable<IdentifierReference> ModuleScopeNonLocalDeclarationReferencesWithinMember
            {
                get
                {
                    if (_moduleScopeNonLocalDeclarationRefs is null)
                    {
                        var member = Member;
                        _moduleScopeNonLocalDeclarationRefs = _moduleDeclarations.AllReferences()
                                .Where(rf => rf.ParentScoping.Equals(member)
                                                && rf.Declaration != member
                                                && !rf.Declaration.DeclarationType.HasFlag(DeclarationType.Parameter)
                                                && !rf.Declaration.IsLocalVariable()
                                                && !rf.Declaration.IsLocalConstant()).ToList();
                    }
                    return _moduleScopeNonLocalDeclarationRefs;
                }
            }
        }
    }
}
