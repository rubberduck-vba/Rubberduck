﻿using System;

namespace Rubberduck.VBEditor.Utility
{
    public interface IDisposalActionContainer<out T>: IDisposable
    {
        T Value { get; }
    }

    internal sealed class DisposalActionContainer<T> : IDisposalActionContainer<T>
    {
        public T Value { get; }
        private readonly Action _disposalAction;

        public DisposalActionContainer(T value, Action disposalAction)
        {
            Value = value;
            _disposalAction = disposalAction;
        }

        private bool _isDisposed = false;
        private readonly object _disposalLockObject = new object();
        public void Dispose()
        {
            lock (_disposalLockObject)
            {
                if (_isDisposed)
                {
                    return;
                }
                _isDisposed = true;
            }

            _disposalAction.Invoke();
        }
    }

    public static class DisposalActionContainer
    {
        public static IDisposalActionContainer<T> Create<T>(T value, Action disposalAction)
        {
            return new DisposalActionContainer<T>(value, disposalAction);
        }
    }
}
