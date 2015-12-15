namespace Rubberduck.Common
{
    public interface IAttachable
    {
        bool IsAttached { get; }

        void Attach();
        void Detach();
    }
}