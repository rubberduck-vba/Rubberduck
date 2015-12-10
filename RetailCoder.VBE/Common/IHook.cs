using System;

namespace Rubberduck.Common
{
    public interface IHook
    {
        event EventHandler<HookEventArgs> MessageReceived;
        void OnMessageReceived();

        bool IsAttached { get; }

        void Attach();
        void Detach();
    }

    public interface ILowLevelKeyboardHook : IHook
    {
        bool EatNextKey { get; set; }
    }
}