using System.Linq;
using Rubberduck.Deployment.Build.Structs;

namespace Rubberduck.Deployment.Build.Writers
{
    public interface IRegistryWriter
    {
        string Write(IOrderedEnumerable<RegistryEntry> entries, string dllName, string tlb32Name, string tlb64Name);
    }
}