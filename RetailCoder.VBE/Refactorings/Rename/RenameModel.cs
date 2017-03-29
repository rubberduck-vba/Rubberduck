using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.PostProcessing;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameModel
    {
        private Rewriters _rewriters;

        private readonly IVBE _vbe;
        public IVBE VBE { get { return _vbe; } }
        
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

        public RenameModel(IVBE vbe, RubberduckParserState state, QualifiedSelection selection, IMessageBox messageBox)
        {
            _vbe = vbe;
            _state = state;
            _declarations = state.AllDeclarations.ToList();
            _selection = selection;
            _messageBox = messageBox;

            _rewriters = new Rewriters(_state);

            AcquireTarget(out _target, Selection);
        }

        public IModuleRewriter GetRewriter(IVBComponent component)
        {
            var qmn = new QualifiedModuleName(component);
            return GetRewriter(qmn);
        }

        public IModuleRewriter GetRewriter(Declaration declaration)
        {
            var qmn = declaration.QualifiedSelection.QualifiedName;
            return GetRewriter(qmn);
        }

        public IModuleRewriter GetRewriter(QualifiedModuleName qmn)
        {
            return _rewriters.GetRewriter(qmn);
        }

        public void Rewrite()
        {
            _rewriters.Rewrite();
        }

        public void ClearRewriters()
        {
            _rewriters = new Rewriters(_state);
        }

        private void AcquireTarget(out Declaration target, QualifiedSelection selection)
        {
            target = _declarations
                .Where(item => item.IsUserDefined && item.DeclarationType != DeclarationType.ModuleOption)
                .FirstOrDefault(item => item.IsSelected(selection) || item.References.Any(r => r.IsSelected(selection)));

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

        internal class Rewriters
        {
            Dictionary<string, IModuleRewriter> _rewriters;
            RubberduckParserState _state;

            public Rewriters(RubberduckParserState state)
            {
                _rewriters = new Dictionary<string, IModuleRewriter>();
                _state = state;
            }

            public IModuleRewriter GetRewriter(QualifiedModuleName qmn)
            {
                IModuleRewriter rewriter;
                if (_rewriters.ContainsKey(qmn.Name))
                {
                    _rewriters.TryGetValue(qmn.Name, out rewriter);
                }
                else
                {
                    rewriter = _state.GetRewriter(qmn);
                    _rewriters.Add(qmn.Name, rewriter);
                }
                return rewriter;
            }

            public void Rewrite()
            {
                foreach(var rewriter in _rewriters.Values)
                {
                    rewriter.Rewrite();
                }
            }
        }

    }
}
