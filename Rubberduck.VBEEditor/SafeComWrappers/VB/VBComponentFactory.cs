using System;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VBA = Rubberduck.VBEditor.SafeComWrappers.VB.VBA;
using VB6 = Rubberduck.VBEditor.SafeComWrappers.VB.VB6;
using VBAIA = Microsoft.Vbe.Interop;
using VB6IA = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB
{
    public static class VBComponentFactory
    {
        public static IVBComponent Create(object vbComponent)
        {
            switch (vbComponent)
            {
                case VBAIA.VBComponent _:
                    return new VBA.VBComponent((VBAIA.VBComponent)vbComponent);

                case VB6IA.VBComponent _:
                    return new VB6.VBComponent((VB6IA.VBComponent)vbComponent);
            }

            throw new NotSupportedException($"VBComponent type '{vbComponent.GetType().FullName}' not supported.");
        }
    }
}
