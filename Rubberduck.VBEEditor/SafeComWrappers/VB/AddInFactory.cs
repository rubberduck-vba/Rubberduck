using System;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VBAIA = Microsoft.Vbe.Interop;
using VB6IA = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB
{
    public static class AddInFactory
    {
        public static IAddIn Create(object application, object addInInst)
        {
            switch (application)
            {
                case VBAIA.VBE _:
                    return new VBA.AddIn((VBAIA.AddIn) addInInst);

                case VB6IA.VBE _:
                    return new VB6.AddIn((VB6IA.AddIn)addInInst);
            }

            throw new NotSupportedException($"Application type '{application.GetType().Name}' not supported.");
        }
    }
}
