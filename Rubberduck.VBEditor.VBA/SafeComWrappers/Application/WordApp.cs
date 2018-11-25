using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class WordApp : HostApplicationBase<Microsoft.Office.Interop.Word.Application>
    {
        public WordApp(IVBE vbe) : base(vbe, "Word", true) { }
    }
}
