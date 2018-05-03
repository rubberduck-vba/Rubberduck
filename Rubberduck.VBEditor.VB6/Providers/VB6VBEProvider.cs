using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop.VB6;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VB6
{
    public class VB6VBEProvider : ISafeComWrapperProvider<IVBE>
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
