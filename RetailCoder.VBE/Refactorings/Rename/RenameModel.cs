using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameModel
    {
        private readonly VBE _vbe;
        public VBE VBE { get { return _vbe; } }
        
        private readonly Declarations _declarations;
        public Declarations Declarations { get { return _declarations; } }

        private Declaration _target;
        public Declaration Target
        {
            get { return _target; }
            set { _target = value; }
        }

        private readonly QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        private readonly VBProjectParseResult _parseResult;
        public VBProjectParseResult ParseResult { get { return _parseResult; } }

        public string NewName { get; set; }

        public RenameModel(VBE vbe, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _vbe = vbe;
            _parseResult = parseResult;
            _declarations = parseResult.Declarations;
            _selection = selection;

            AcquireTarget(out _target, Selection);
        }

        private void AcquireTarget(out Declaration target, QualifiedSelection selection)
        {
            target = _declarations.Items
                .Where(item => !item.IsBuiltIn && item.DeclarationType != DeclarationType.ModuleOption)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item)
                                      || IsSelectedReference(selection, item));

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

            var confirm = MessageBox.Show(message, RubberduckUI.RenameDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (confirm == DialogResult.No)
            {
                target = null;
                return;
            }

            target = interfaceMember;
        }

        private bool IsSelectedReference(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.References.Any(r =>
                r.QualifiedModuleName.Project == selection.QualifiedName.Project
                && r.QualifiedModuleName.ComponentName == selection.QualifiedName.ComponentName
                && r.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName.Project == selection.QualifiedName.Project
                   && declaration.QualifiedName.QualifiedModuleName.ComponentName == selection.QualifiedName.ComponentName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
        }
    }
}
