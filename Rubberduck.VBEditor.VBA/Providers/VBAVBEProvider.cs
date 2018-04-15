using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBAVBEProvider : ISafeComWrapperProvider<IVBE>
    {
        public bool CanProvideFor(object comObject)
        {
            return comObject is VB.VBE;
        }

        public IVBE Provide(object comObject)
        {
            return new VBE((VB.VBE)comObject);
        }
    }
}
