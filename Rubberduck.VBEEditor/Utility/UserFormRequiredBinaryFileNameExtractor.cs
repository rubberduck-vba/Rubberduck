using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;
using Path = System.IO.Path;

namespace Rubberduck.VBEditor.Utility
{
    public class UserFormRequiredBinaryFileNameExtractor : IRequiredBinaryFilesFromFileNameExtractor
    {
        public ICollection<ComponentType> SupportedComponentTypes => new List<ComponentType>{ComponentType.UserForm};

        public ICollection<string> RequiredBinaryFiles(string fileName, ComponentType componentType)
        {
            if (!File.Exists(fileName))
            {
                return null;
            }

            if (!SupportedComponentTypes.Contains(componentType))
            {
                return null;
            }

            if (componentType.FileExtension() != Path.GetExtension(fileName))
            {
                return null;
            }

            var regExPattern = "OleObjectBlob\\s+=\\s+\"([^\"]+)\":";
            var regEx = new Regex(regExPattern);
            var contents = File.ReadLines(fileName, Encoding.Default);
            
            foreach(var codeLine in contents)
            {
                var match = regEx.Match(codeLine);
                if (match.Success)
                {
                    return new List<string>{match.Groups[1].Value};
                }
            }

            var fallbackBinaryName = Path.GetFileNameWithoutExtension(fileName) + componentType.BinaryFileExtension();
            return new List<string>{fallbackBinaryName};
        }
    }
}