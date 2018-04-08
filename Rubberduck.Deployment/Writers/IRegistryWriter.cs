using System.Linq;
using Rubberduck.Deployment.Structs;

namespace Rubberduck.Deployment.Writers
{
    public interface IRegistryWriter
    {
        string Write(IOrderedEnumerable<RegistryEntry> entries);
    }
}