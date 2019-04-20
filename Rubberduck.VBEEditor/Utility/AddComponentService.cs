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
            IComponentSourceCodeHandler codePaneComponentSourceCodeProvider,
            IComponentSourceCodeHandler attributesComponentSourceCodeProvider)
        {
            _projectsProvider = projectsProvider;
            _codePaneSourceCodeHandler = codePaneComponentSourceCodeProvider;
            _attributeSourceCodeHandler = attributesComponentSourceCodeProvider;
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
                    using (var loadedComponent = sourceCodeHandler.SubstituteCode(newComponent, code))
                    {
                        AddPrefix(loadedComponent, prefixInModule);
                        ShowComponent(loadedComponent);
                    }
                }
                else
                {
                    AddPrefix(newComponent, prefixInModule);
                    ShowComponent(newComponent);
                }
            }
        }

        private static void ShowComponent(IVBComponent component)
        {
            if (component == null)
            {
                return;
            }

            using (var codeModule = component.CodeModule)
            {
                if (codeModule == null)
                {
                    return;
                }

                using (var codePane = codeModule.CodePane)
                {
                    codePane.Show();
                }
            }
        }

        private static void AddPrefix(IVBComponent module, string prefix)
        {
            if (prefix == null || module == null)
            {
                return;
            }

            using (var codeModule = module.CodeModule)
            {
                codeModule.InsertLines(1, prefix);
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