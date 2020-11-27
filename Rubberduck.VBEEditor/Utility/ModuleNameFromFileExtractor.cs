using System.Collections.Generic;
using System.IO.Abstractions;
using System.Linq;
using System.Text;
using Rubberduck.InternalApi.Common;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.Utility
{
    public interface IModuleNameFromFileExtractor
    {
        string ModuleName(string filename);
    }

    public class ModuleNameFromFileExtractor : IModuleNameFromFileExtractor
    {
        private readonly IFileSystem _fileSystem = FileSystemProvider.FileSystem;

        public string ModuleName(string filename)
        {
            if (!_fileSystem.File.Exists(filename))
            {
                return null;
            }

            if (!SupportedExtensions.Contains(_fileSystem.Path.GetExtension(filename)))
            {
                return null;
            }

            var contents = _fileSystem.File.ReadLines(filename, Encoding.Default);
            var nameLine = contents.FirstOrDefault(line => line.StartsWith("Attribute VB_Name = "));
            if (nameLine == null)
            {
                return _fileSystem.Path.GetFileNameWithoutExtension(filename);
            }

            //The format is Attribute VB_Name = "ModuleName"
            return nameLine.Substring("Attribute VB_Name = ".Length + 1, nameLine.Length - "Attribute VB_Name = ".Length - 2);
        }

        private static ICollection<string> SupportedExtensions => 
            ComponentTypeExtensions.ComponentTypesForExtension(VBEKind.Hosted).Keys
            .Concat(ComponentTypeExtensions.ComponentTypesForExtension(VBEKind.Standalone).Keys)
            .ToHashSet();
    }
}