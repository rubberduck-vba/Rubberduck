using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SourceCodeHandling;

namespace Rubberduck.VBEditor.Utility
{
    public class AddComponentService : IAddComponentService
    {
        private readonly IProjectsProvider _projectsProvider;
        private readonly IComponentSourceCodeHandler _codePaneSourceCodeHandler;
        private readonly IComponentSourceCodeHandler _attributeSourceCodeHandler;

        public AddComponentService(
            IProjectsProvider projectsProvider,
            IComponentSourceCodeHandler codePaneSourceCodeHandler,
            IComponentSourceCodeHandler attributeSourceCodeHandler)
        {
            _projectsProvider = projectsProvider;
            _codePaneSourceCodeHandler = codePaneSourceCodeHandler;
            _attributeSourceCodeHandler = attributeSourceCodeHandler;
        }

        public void AddComponent(string projectId, ComponentType componentType, string code = null, string additionalPrefixInModule = null)
        {
            AddComponent(_codePaneSourceCodeHandler, projectId, componentType, code, additionalPrefixInModule);
        }

        public void AddComponentWithAttributes(string projectId, ComponentType componentType, string code, string prefixInModule = null)
        {
            AddComponent(_attributeSourceCodeHandler, projectId, componentType, code, prefixInModule);
        }

        public void AddComponent(IComponentSourceCodeHandler sourceCodeHandler, string projectId, ComponentType componentType, string code = null, string prefixInModule = null)
        {
            using (var newComponent = CreateComponent(projectId, componentType))
            {
                if (newComponent == null)
                {
                    return;
                }
                
                if (code != null)
                {
                    sourceCodeHandler.SubstituteCode(newComponent, code);
                }

                if (prefixInModule != null)
                {
                    using (var codeModule = newComponent.CodeModule)
                    {
                        codeModule.InsertLines(1, prefixInModule);
                    }
                }
            }
        }

        private IVBComponent CreateComponent(string projectId, ComponentType componentType)
        {
            var componentsCollection = _projectsProvider.ComponentsCollection(projectId);
            if (componentsCollection == null)
            {
                return null;
            }

            return componentsCollection.Add(componentType);
        }
    }
}