using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.VBEditor.Application
{
    public class HostDocumentMacro
    {
        private readonly IEnumerable<string> _eventNames;

        internal HostDocumentMacro(int id, string controlName, string progId, IEnumerable<string> eventNames)
            : this(id, controlName, string.Empty)
        {
            ProgId = progId;
            _eventNames = eventNames;
        }

        internal HostDocumentMacro(int id, string controlName, string macroName)
        {
            Id = id;
            ControlName = controlName;
            MacroName = macroName;
            _eventNames = Enumerable.Empty<string>();
        }

        public int Id { get; }
        public string ProgId { get; }
        public string ControlName { get; }
        public string MacroName { get; }
        public IEnumerable<string> Handlers { get { return _eventNames.Select(e => $"{ControlName}_{e}"); } }
    }
}