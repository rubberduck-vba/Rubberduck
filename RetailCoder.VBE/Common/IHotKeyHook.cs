namespace Rubberduck.Common
{
    public interface IHotKeyHook : IHook
    {
        HookInfo HookInfo { get; }
        bool IsTwoStepHotKey { get; }
    }
}