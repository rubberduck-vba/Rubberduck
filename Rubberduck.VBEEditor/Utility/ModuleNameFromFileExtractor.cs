using System.IO;
using System.Linq;
using System.Text;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.VBEditor.Utility
{
    public interface IModuleNameFromFileExtractor
    {
        string ModuleName(string filename);
    }

    public class ModuleNameFromFileExtractor : IModuleNameFromFileExtractor
    {
        public string ModuleName(string filename)
        {
            if (!File.Exists(filename))
            {
                return null;
            }

            //We cannot read binary files.
            if(ComponentTypeExtensions.FormBinaryExtension.Equals(Path.GetExtension(filename)))
            {
                return Path.GetFileNameWithoutExtension(filename);
            }

            var contents = File.ReadLines(filename, Encoding.Default);
            var nameLine = contents.FirstOrDefault(line => line.StartsWith("Attribute VB_Name = "));
            if (nameLine == null)
            {
                return Path.GetFileNameWithoutExtension(filename);
            }

            //The format is Attribute VB_Name = "ModuleName"
            return nameLine.Substring("Attribute VB_Name = ".Length + 1, nameLine.Length - "Attribute VB_Name = ".Length - 2);
        }
    }
}