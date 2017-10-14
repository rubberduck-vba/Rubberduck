using System;

namespace Rubberduck.VBEditor.SafeComWrappers.VB.Abstract
{
    internal interface IVBComIds
    {
        Guid VBComponentEvents { get; }
        Guid VBProjectEvents { get; }
        IVBComponentEventDispIds VBComponent { get; }                
        IVBProjectEventDispIds VBProject { get; }
    }

    internal interface IVBComponentEventDispIds
    {
        int Added { get; }
        int Removed { get; }
        int Renamed { get; }
        int Selected { get; }
        int Activated { get; }
        int Reloaded { get; }
    }

    internal interface IVBProjectEventDispIds
    {
        int Added { get; }
        int Removed { get; }
        int Renamed { get; }
        int Activated { get; }
    }
}
