namespace Rubberduck.UI.Command
{
    public interface ICommand
    {
        void Execute();
    }

    public interface ICommand<in T> : ICommand
    {
        void Execute(T parameter);
    }
}
