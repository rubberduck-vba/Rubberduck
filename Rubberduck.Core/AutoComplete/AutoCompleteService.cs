using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteService : IAutoCompleteService, IDisposable
    {
        private readonly IReadOnlyList<IAutoComplete> _autoCompletions;

        public AutoCompleteService(IReadOnlyList<IAutoComplete> autoCompletes)
        {
            _autoCompletions = autoCompletes;
            VBENativeServices.CaretHidden += VBENativeServices_CaretHidden;
        }

        public event EventHandler AutoCompleteTriggered;

        private QualifiedSelection? _lastSelection;
        private string _lastCode;
        string _contentHash;

        private void VBENativeServices_CaretHidden(object sender, AutoCompleteEventArgs e)
        {
            AutoCompleteTriggered?.Invoke(this, e);
            var selection = e.CodePane.Selection;
            var qualifiedSelection = e.CodePane.GetQualifiedSelection();

            if (!selection.IsSingleCharacter || e.OldCode.Equals(_lastCode) || qualifiedSelection.Value.Equals(_lastSelection) || string.IsNullOrWhiteSpace(e.OldCode) || e.ContentHash == _contentHash)
            {
                return;
            }

            foreach (var autoCompletion in _autoCompletions.Where(auto => auto.IsEnabled))
            {
                if (autoCompletion.Execute(e))
                {
                    _lastSelection = qualifiedSelection;
                    _lastCode = e.NewCode;
                    using (var module = e.CodePane.CodeModule)
                    {
                        _contentHash = module.ContentHash();
                    }
                    break;
                }
            }
        }

        public void Dispose()
        {
            VBENativeServices.CaretHidden -= VBENativeServices_CaretHidden;
        }
    }
}
