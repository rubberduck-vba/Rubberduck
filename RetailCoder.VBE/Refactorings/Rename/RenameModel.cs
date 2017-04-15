using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameModel
    {
        public IVBE VBE { get; }

        private readonly IList<Declaration> _declarations;
        public IEnumerable<Declaration> Declarations => _declarations;

        private Declaration _target;
        public Declaration Target
        {
            get { return _target; }
            set { _target = value; }
        }

        public QualifiedSelection Selection { get; }

        public RubberduckParserState State { get; }

        public string NewName { get; set; }

        public RenameModel(IVBE vbe, RubberduckParserState state, QualifiedSelection selection)
        {
            VBE = vbe;
            State = state;
            _declarations = state.AllDeclarations.ToList();
            Selection = selection;

            AcquireTarget(out _target, Selection);
        }

        private void AcquireTarget(out Declaration target, QualifiedSelection selection)
        {
            target = _declarations
                .Where(item => item.IsUserDefined && item.DeclarationType != DeclarationType.ModuleOption)
                .FirstOrDefault(item => item.IsSelected(selection) || item.References.Any(r => r.IsSelected(selection)));
        }
    }
}
