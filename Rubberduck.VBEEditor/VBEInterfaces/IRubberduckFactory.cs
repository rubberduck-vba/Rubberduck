namespace Rubberduck.VBEditor.VBEInterfaces
{
    public interface IRubberduckFactory<out TPresenter>
    {
        TPresenter Create(object obj);
    }
}