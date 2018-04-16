using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor
{
    public interface ISafeComWrapperProvider<out TWrapper> where TWrapper : ISafeComWrapper
    {
        bool CanProvideFor(object comObject);

        TWrapper Provide(object comObject);
    }
}
