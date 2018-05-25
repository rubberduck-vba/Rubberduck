using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteService : IAutoCompleteService, IDisposable
    {
        public event EventHandler TypingCode;
        private readonly IReadOnlyList<IAutoComplete> _autocompletions;

        public AutoCompleteService(IReadOnlyList<IAutoComplete> autoCompletes)
        {
            _autocompletions = autoCompletes;
            VBENativeServices.TypingCode += VBENativeServices_TypingCode;
        }

        QualifiedSelection? _lastSelection;
        string _lastCode;

        private void VBENativeServices_TypingCode(object sender, AutoCompleteEventArgs e)
        {
            TypingCode?.Invoke(this, e);
            var selection = e.CodePane.Selection;
            var qualifiedSelection = e.CodePane.GetQualifiedSelection();

            if (!selection.IsSingleCharacter || e.OldCode.Equals(_lastCode) || qualifiedSelection.Value.Equals(_lastSelection))
            {
                return;
            }

            foreach (var autocomplete in _autocompletions.Where(auto => auto.IsEnabled))
            {
                if (autocomplete.Execute(e))
                {
                    _lastSelection = qualifiedSelection;
                    _lastCode = e.NewCode;
                    break;
                }
            }
        }

        public void Dispose()
        {
            VBENativeServices.TypingCode -= VBENativeServices_TypingCode;
        }
    }
}
