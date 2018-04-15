using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using VB = Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class VBAAddInProvider : ISafeComWrapperProvider<IAddIn>
    {
        public bool CanProvideFor(object comObject)
        {
            return comObject is VB.AddIn;
        }

        public IAddIn Provide(object comObject)
        {
            return new AddIn((VB.AddIn)comObject);
        }
    }
}
