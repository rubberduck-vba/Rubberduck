﻿using System;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.VBEditor.WindowsApi
{
    public interface IWindowEventProvider : IDisposable
    {
        event EventHandler<WindowChangedEventArgs> FocusChange;
    }
}
