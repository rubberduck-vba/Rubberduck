using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Extensions;

namespace Rubberduck.VBEHost
{
    public interface IVBEAdapter
    {
        IHostApplication HostApplication { get; }
        IEnumerable<ICommandBarAdapter> CommandBars { get; }
        IEnumerable<IVBProjectAdapter> Projects { get; }

    }

    public interface ICommandBarAdapter
    {
        
    }

    public interface IVBProjectAdapter
    {
        string Path { get; set; }
        string Name { get; set; }
        IEnumerable<IVBComponentAdapter> Components { get; }
        void AddComponent();
        void RemoveComponent(IVBComponentAdapter component);
    }

    public interface IVBComponentAdapter
    {
        string Path { get; set; }
        string Name { get; set; }
        ICodePaneAdapter CodePane { get; }
    }

    public interface ICodePaneAdapter
    {
        string Code { get; }
    }
}
