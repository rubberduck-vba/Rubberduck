using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameModel
    {
        private readonly VBE _vbe;
        public VBE VBE { get { return _vbe; } }
        
        private readonly IList<Declaration> _declarations;
        public IEnumerable<Declaration> Declarations { get { return _declarations; } }

        private Declaration _target;
        public Declaration Target
        {
            get { return _target; }
            set { _target = value; }
        }

        private readonly QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        private readonly RubberduckParserState _state;
        public RubberduckParserState State { get { return _state; } }

        public string NewName { get; set; }

        private readonly IMessageBox _messageBox;

        public RenameModel(VBE vbe, RubberduckParserState state, QualifiedSelection selection, IMessageBox messageBox)
        {
            _vbe = vbe;
            _state = state;
            _declarations = state.AllDeclarations.ToList();
            _selection = selection;
            _messageBox = messageBox;

            AcquireTarget(out _target, Selection);
        }

        private void AcquireTarget(out Declaration target, QualifiedSelection selection)
        {
            target = _declarations
                .Where(item => !item.IsBuiltIn && item.DeclarationType != DeclarationType.ModuleOption)
                .FirstOrDefault(item => item.IsSelected(selection)
                                      || item.References.Any(r => r.IsSelected(selection)));

            PromptIfTargetImplementsInterface(ref target);
        }

        public void PromptIfTargetImplementsInterface(ref Declaration target)
        {
            var declaration = target;
            var interfaceImplementation = _declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (target == null || interfaceImplementation == null)
            {
                return;
            }

            var interfaceMember = _declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.RenamePresenter_TargetIsInterfaceMemberImplementation, target.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = _messageBox.Show(message, RubberduckUI.RenameDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirm == DialogResult.No)
            {
                target = null;
                return;
            }

            target = interfaceMember;
        }
    }
}
