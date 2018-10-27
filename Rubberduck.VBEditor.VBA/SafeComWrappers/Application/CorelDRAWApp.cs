using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class CorelDRAWApp : HostApplicationBase<Corel.GraphicsSuite.Interop.CorelDRAW.Application>
    {
        public CorelDRAWApp(IVBE vbe) : base(vbe, "CorelDRAW", true) { }

		//TODO:Can only get a CorelDraw application if at least one document is open in CorelDraw.
    }
}
