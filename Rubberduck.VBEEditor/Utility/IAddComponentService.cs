using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.Utility
{
    public interface IAddComponentService
    {
        void AddComponent(string projectId, ComponentType componentType, string code = null, string additionalPrefixInModule = null);
        void AddComponentWithAttributes(string projectId, ComponentType componentType, string code, string prefixInModule = null);
    }
}