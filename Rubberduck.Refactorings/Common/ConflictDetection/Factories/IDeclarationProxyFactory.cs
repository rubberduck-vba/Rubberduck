using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IDeclarationProxyFactory
    {
        IConflictDetectionDeclarationProxy CreateProxy(Declaration prototype);
        IConflictDetectionDeclarationProxy CreateNewEntityProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, IConflictDetectionDeclarationProxy parentProxy);
        IConflictDetectionDeclarationProxy CreateNewEntityProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, QualifiedModuleName qmn);
        IConflictDetectionModuleDeclarationProxy CreateProxyNewModule(string projectID, ComponentType componentType, string proposedName);
    }

    public class ConflictDetectionDeclarationProxyFactory : IDeclarationProxyFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public ConflictDetectionDeclarationProxyFactory(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public IConflictDetectionDeclarationProxy CreateProxy(Declaration prototype)
        {
            if (prototype is ModuleDeclaration module)
            {
                return new ModuleConflictDetectionDeclarationProxy(module);
            }

            module = _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(prototype.QualifiedModuleName) as ModuleDeclaration;
            return new ConflictDetectionDeclarationProxy(prototype, module);
        }

        public IConflictDetectionDeclarationProxy CreateNewEntityProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, IConflictDetectionDeclarationProxy parentProxy)
        {
            switch (declarationType)
            {
                case DeclarationType.UserDefinedType:
                    return new UDTConflictDetectionDeclarationProxy(identifier, declarationType, accessibility, parentProxy);
                case DeclarationType.UserDefinedTypeMember:
                    return new UDTMemberConflictDetectionDeclarationProxy(identifier, declarationType, accessibility, parentProxy);
                case DeclarationType.Enumeration:
                    return new EnumConflictDetectionDeclarationProxy(identifier, declarationType, accessibility, parentProxy);
                case DeclarationType.EnumerationMember:
                    return new EnumMemberConflictDetectionDeclarationProxy(identifier, declarationType, accessibility, parentProxy);
                default:
                    var proxy = new ConflictDetectionDeclarationProxy(identifier, declarationType, accessibility, parentProxy.Prototype)
                    {
                        ParentProxy = parentProxy,
                        TargetModule = parentProxy.TargetModule
                    };
                    return proxy;
            }
        }

        public IConflictDetectionDeclarationProxy CreateNewEntityProxy(string identifier, DeclarationType declarationType, Accessibility accessibility, QualifiedModuleName qmn)
        {
            var module = _declarationFinderProvider.DeclarationFinder.ModuleDeclaration(qmn);
            var parentProxy = CreateProxy(module);
            return CreateNewEntityProxy(identifier, declarationType, accessibility, parentProxy);
        }

        public IConflictDetectionModuleDeclarationProxy CreateProxyNewModule(string projectId, ComponentType componentType, string proposedName)
        {
            var project = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Project)
                .SingleOrDefault(p => p.ProjectId == projectId);
            return new ModuleConflictDetectionDeclarationProxy(project, componentType, proposedName);
        }
    }
}
