using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.Abstract
{
    internal interface IComIds
    {
        Guid VBComponentsEventsGuid { get; }
        Guid VBProjectsEventsGuid { get; }
        IComponentEventDispIds ComponentEventDispIds { get; }                
        IProjectEventDispIds ProjectEventDispIds { get; }
    }

    internal interface IComponentEventDispIds
    {
        int ItemAdded { get; }
        int ItemRemoved { get; }
        int ItemRenamed { get; }
        int ItemSelected { get; }
        int ItemActivated { get; }
        int ItemReloaded { get; }
    }

    internal interface IProjectEventDispIds
    {
        int ItemAdded { get; }
        int ItemRemoved { get; }
        int ItemRenamed { get; }
        int ItemActivated { get; }
    }
}
