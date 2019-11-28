using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.Utility
{
    public interface IRequiredBinaryFilesFromFileNameExtractor
    {
        ICollection<ComponentType> SupportedComponentTypes { get; }
        ICollection<string> RequiredBinaryFiles(string fileName, ComponentType componentType);
    }
}