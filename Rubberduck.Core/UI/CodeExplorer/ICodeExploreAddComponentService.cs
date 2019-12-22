using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer
{
    public interface ICodeExplorerAddComponentService
    {
        void AddComponent(CodeExplorerItemViewModel node, ComponentType componentType, string code = null);
        void AddComponentWithAttributes(CodeExplorerItemViewModel node, ComponentType componentType, string code, string additionalPrefixInModule = null);
    }
}