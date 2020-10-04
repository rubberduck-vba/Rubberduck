using System.Collections.Generic;
using System.IO.Abstractions;
using System.Text;
using System.Text.RegularExpressions;
using Rubberduck.InternalApi.Common;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.VBEditor.Utility
{
    public class UserFormRequiredBinaryFileNameExtractor : IRequiredBinaryFilesFromFileNameExtractor
    {
        private readonly IFileSystem _fileSystem = FileSystemProvider.FileSystem;

        public ICollection<ComponentType> SupportedComponentTypes => new List<ComponentType>{ComponentType.UserForm};

        public ICollection<string> RequiredBinaryFiles(string fileName, ComponentType componentType)
        {
            if (!_fileSystem.File.Exists(fileName))
            {
                return null;
            }

            if (!SupportedComponentTypes.Contains(componentType))
            {
                return null;
            }

            if (componentType.FileExtension() != _fileSystem.Path.GetExtension(fileName))
            {
                return null;
            }

            var regExPattern = "OleObjectBlob\\s+=\\s+\"([^\"]+)\":";
            var regEx = new Regex(regExPattern);
            var contents = _fileSystem.File.ReadLines(fileName, Encoding.Default);
            
            foreach(var codeLine in contents)
            {
                var match = regEx.Match(codeLine);
                if (match.Success)
                {
                    return new List<string>{match.Groups[1].Value};
                }
            }

            var fallbackBinaryName = _fileSystem.Path.GetFileNameWithoutExtension(fileName) + componentType.BinaryFileExtension();
            return new List<string>{fallbackBinaryName};
        }
    }
}