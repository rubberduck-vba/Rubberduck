using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NLog;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public class MemberAttributeRecoverer : IMemberAttributeRecovererWithSettableRewritingManager
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IParseManager _parseManager;
        private readonly IAttributesUpdater _attributesUpdater;
        private IRewritingManager _rewritingManager;

        private readonly
            IDictionary<QualifiedModuleName, IDictionary<string, HashSet<AttributeNode>>> _attributesToRecover
                = new Dictionary<QualifiedModuleName, IDictionary<string, HashSet<AttributeNode>>>();

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public MemberAttributeRecoverer(IDeclarationFinderProvider declarationFinderProvider,
            IParseManager parseManager, IAttributesUpdater attributesUpdater)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _parseManager = parseManager;
            _attributesUpdater = attributesUpdater;
        }

        //This needs to be property injected because this class will be constructor injected into the RewritingManager that needs to set itself as the dependency here.
        public IRewritingManager RewritingManager
        {
            set => _rewritingManager = value;
        }

        public void RecoverCurrentMemberAttributesAfterNextParse(IEnumerable<QualifiedMemberName> members)
        {
            var declarationsByModule = MemberDeclarationsByModule(members);
            RecoverCurrentMemberAttributesAfterNextParse(declarationsByModule);
        }

        private IDictionary<QualifiedModuleName, IEnumerable<Declaration>> MemberDeclarationsByModule(IEnumerable<QualifiedMemberName> members)
        {
            var membersByModule = members.GroupBy(member => member.QualifiedModuleName)
                .ToDictionary(group => group.Key, group => group.ToHashSet());
            var declarationFinder = _declarationFinderProvider.DeclarationFinder;
            var memberDeclarationsPerModule = new Dictionary<QualifiedModuleName, IEnumerable<Declaration>>();
            foreach (var module in membersByModule.Keys)
            {
                var memberDeclarationsInModule = declarationFinder.Members(module)
                    .Where(declaration => membersByModule[module].Contains(declaration.QualifiedName)
                                          && declaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module));
                memberDeclarationsPerModule.Add(module, memberDeclarationsInModule);
            }

            return memberDeclarationsPerModule;
        }

        private void RecoverCurrentMemberAttributesAfterNextParse(
            IDictionary<QualifiedModuleName, IEnumerable<Declaration>> declarationsByModule)
        {
            SaveAttributesToRecover(declarationsByModule);

            if (_attributesToRecover.Any())
            {
                PrimeRecoveryOfAttributes();
            }
        }

        private void SaveAttributesToRecover(IDictionary<QualifiedModuleName, IEnumerable<Declaration>> declarationsByModule)
        {
            _attributesToRecover.Clear();
            foreach (var module in declarationsByModule.Keys)
            {
                var attributesByMember = declarationsByModule[module]
                    .Where(decl => decl.Attributes.Any())
                    .ToDictionary(
                        decl => decl.IdentifierName,
                        decl => (HashSet<AttributeNode>)decl.Attributes);
                _attributesToRecover.Add(module, attributesByMember);
            }
        }

        private void PrimeRecoveryOfAttributes()
        {
            _parseManager.StateChanged += ExecuteRecoveryOfAttributes;
        }

        private void ExecuteRecoveryOfAttributes(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.ResolvedDeclarations)
            {
                return;
            }

            StopRecoveringAttributesOnNextParse();
            
            var rewriteSession = _rewritingManager.CheckOutAttributesSession();
            foreach (var module in _attributesToRecover.Keys)
            {
                RecoverAttributes(rewriteSession, module, _attributesToRecover[module]);
            }

            if (!rewriteSession.CheckedOutModules.Any())
            {
                //There is nothing we can do.
                return;
            }

            CancelTheCurrentParse();

            Task.Run(() => rewriteSession.TryRewrite());

            EndTheCurrentParse(e.Token);
        }

        private void StopRecoveringAttributesOnNextParse()
        {
            _parseManager.StateChanged -= ExecuteRecoveryOfAttributes;
        }

        private void CancelTheCurrentParse()
        {
            _parseManager.OnParseCancellationRequested(this);
        }

        private void RecoverAttributes(IRewriteSession rewriteSession, QualifiedModuleName module, IDictionary<string, HashSet<AttributeNode>> attributesByMember)
        {
            var membersWithAttributesToRecover = attributesByMember.Keys.ToHashSet();
            var declarationFinder = _declarationFinderProvider.DeclarationFinder;
            var declarationsWithAttributesToRecover = declarationFinder.Members(module)
                .Where(decl => membersWithAttributesToRecover.Contains(decl.IdentifierName) 
                               && decl.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
                .ToList();

            if (membersWithAttributesToRecover.Count != declarationsWithAttributesToRecover.Count)
            {
                LogFailureToRecoverAllAttributes(module, membersWithAttributesToRecover, declarationsWithAttributesToRecover);
            }

            foreach (var declaration in declarationsWithAttributesToRecover)
            {
                RecoverAttributes(rewriteSession, declaration, attributesByMember[declaration.IdentifierName]);
            }
        }

        private void LogFailureToRecoverAllAttributes(QualifiedModuleName module, IEnumerable<string> membersWithAttributesToRecover, List<Declaration> declarationsWithAttributesToRecover)
        {
            _logger.Warn("Could not recover the attributes for all members because one or more members could no longer be found.");

            var membersWithoutDeclarations = membersWithAttributesToRecover.ToHashSet();
            membersWithoutDeclarations.ExceptWith(declarationsWithAttributesToRecover.Select(decl => decl.IdentifierName));
            foreach (var member in membersWithoutDeclarations)
            {
                _logger.Trace($"Could not recover the attributes for member {member} in module {module} because a member of that name exists no longer.");
            }
        }

        private void RecoverAttributes(IRewriteSession rewriteSession, Declaration declaration, IEnumerable<AttributeNode> attributes)
        {
            foreach (var attribute in attributes)
            {
                if (!declaration.Attributes.Contains(attribute))
                {
                    _attributesUpdater.AddAttribute(rewriteSession, declaration, attribute.Name, attribute.Values);
                }
            }
        }

        private void EndTheCurrentParse(CancellationToken currentCancellationToken)
        {
            currentCancellationToken.ThrowIfCancellationRequested();
        }

        public void RecoverCurrentMemberAttributesAfterNextParse(IEnumerable<QualifiedModuleName> modules)
        {
            var declarationsByModule = MemberDeclarationsByModule(modules);
            RecoverCurrentMemberAttributesAfterNextParse(declarationsByModule);
        }

        private IDictionary<QualifiedModuleName, IEnumerable<Declaration>> MemberDeclarationsByModule(IEnumerable<QualifiedModuleName> modules)
        {
            var declarationFinder = _declarationFinderProvider.DeclarationFinder;
            var declarationsByModule = modules.ToDictionary(
                module => module, 
                module => declarationFinder.Members(module)
                    .Where(decl => !decl.DeclarationType.HasFlag(DeclarationType.Module)
                                   && decl.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module)));
            return declarationsByModule;
        }
    }
}