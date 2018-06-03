using System;

namespace Rubberduck.AutoComplete
{
    public interface IAutoCompleteService
    {
        event EventHandler AutoCompleteTriggered;
    }
}
