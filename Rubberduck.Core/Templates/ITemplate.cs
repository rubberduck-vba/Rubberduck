namespace Rubberduck.Templates
{
    public interface ITemplate
    {
        string Name { get; }
        bool IsUserDefined { get; }
        string Caption { get; }
        string Description { get; }
        string Read();
        void Write(string content);
    }
}