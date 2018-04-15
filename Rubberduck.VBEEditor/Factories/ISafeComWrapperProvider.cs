namespace Rubberduck.VBEditor
{
    public interface ISafeComWrapperProvider<out TWrapper>
    {
        bool CanProvideFor(object comObject);

        TWrapper Provide(object comObject);
    }
}
