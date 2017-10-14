using System;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using VBAIA = Microsoft.Vbe.Interop;
using VB6IA = Microsoft.VB6.Interop.VBIDE;

namespace Rubberduck.VBEditor.SafeComWrappers.VB
{
    public static class VBProjectFactory
    {
        public static IVBProject Create(object project)
        {
            switch (project)
            {
                case VBAIA.VBProject vbap:
                    return new VBA.VBProject(vbap);

                case VB6IA.VBProject vb6p:
                    return new VB6.VBProject(vb6p);
            }

            throw new NotSupportedException($"Project type '{project.GetType().FullName}' not supported.");
        }
    }
}