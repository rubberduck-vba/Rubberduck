using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteService : IDisposable
    {
        private readonly IReadOnlyList<IAutoComplete> _autoCompletes;
        private QualifiedSelection? _lastSelection;
        private string _lastCode;
        private string _contentHash;

        public AutoCompleteService(IReadOnlyList<IAutoComplete> autoCompletes)
        {
            _autoCompletes = autoCompletes;
            VBENativeServices.CaretHidden += VBENativeServices_CaretHidden;
        }

        private void VBENativeServices_CaretHidden(object sender, AutoCompleteEventArgs e)
        {
            if (e.ContentHash == _contentHash)
            {
                return;
            }

            var qualifiedSelection = e.CodePane.GetQualifiedSelection();
            var selection = qualifiedSelection.Value.Selection;

            foreach (var autoComplete in _autoCompletes.Where(auto => auto.IsEnabled))
            {
                if (autoComplete.Execute(e))
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
