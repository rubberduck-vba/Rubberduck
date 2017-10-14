using System;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VBAIA = Microsoft.Vbe.Interop;
using VB6IA = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB
{
    public static class VBEFactory
    {
        public static IVBE Create(object application)
        {
            switch (application)
            {
                case VBAIA.VBE vbae:
                    return new VBA.VBE(vbae);   
                    
                case VB6IA.VBE vb6e:
                    return new VB6.VBE(vb6e);
            }

            throw new NotSupportedException($"Application type '{application.GetType().FullName}' not supported.");
        }
    }
}
