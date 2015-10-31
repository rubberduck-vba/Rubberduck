using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersModel
    {
        private readonly RubberduckParserState _parseResult;
        public RubberduckParserState ParseResult { get { return _parseResult; } }

        private readonly IList<Declaration> _declarations;
        public IEnumerable<Declaration> Declarations { get { return _declarations; } }

        public Declaration TargetDeclaration { get; private set; }
        public List<Parameter> Parameters { get; set; }

        private readonly IMessageBox _messageBox;

        public RemoveParametersModel(RubberduckParserState parseResult, QualifiedSelection selection, IMessageBox messageBox)
        {
            _parseResult = parseResult;
            _declarations = parseResult.AllDeclarations.ToList();
            _messageBox = messageBox;

            AcquireTarget(selection);

            Parameters = new List<Parameter>();
            LoadParameters();
        }

        private void AcquireTarget(QualifiedSelection selection)
        {
            TargetDeclaration = Declarations.FindSelection(selection, ValidDeclarationTypes);
            TargetDeclaration = PromptIfTargetImplementsInterface();
            TargetDeclaration = GetGetter();
        }

        private void LoadParameters()
        {
            if (TargetDeclaration == null) { return; }

            Parameters.Clear();

            var index = 0;
            Parameters = GetParameters().Select(arg => new Parameter(arg, index++)).ToList();

            if (TargetDeclaration.DeclarationType == DeclarationType.PropertyLet ||
                TargetDeclaration.DeclarationType == DeclarationType.PropertySet)
            {
                Parameters.Remove(Parameters.Last());
            }
        }

        private IEnumerable<Declaration> GetParameters()
        {
            var targetSelection = new Selection(TargetDeclaration.Context.Start.Line,
                TargetDeclaration.Context.Start.Column,
                TargetDeclaration.Context.Stop.Line,
                TargetDeclaration.Context.Stop.Column);

            return Declarations.Where(d => d.DeclarationType == DeclarationType.Parameter
                                       && d.ComponentName == TargetDeclaration.ComponentName
                                       && d.Project.Equals(TargetDeclaration.Project)
                                       && targetSelection.Contains(d.Selection))
                              .OrderBy(item => item.Selection.StartLine)
                              .ThenBy(item => item.Selection.StartColumn);
        }

        public static readonly DeclarationType[] ValidDeclarationTypes =
        {
            DeclarationType.Event,
            DeclarationType.Function,
            DeclarationType.Procedure,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };

        private Declaration PromptIfTargetImplementsInterface()
        {
            var declaration = TargetDeclaration;
            var interfaceImplementation = Declarations.FindInterfaceImplementationMembers().SingleOrDefault(m => m.Equals(declaration));
            if (declaration == null || interfaceImplementation == null)
            {
                return declaration;
            }

            var interfaceMember = Declarations.FindInterfaceMember(interfaceImplementation);
            var message = string.Format(RubberduckUI.Refactoring_TargetIsInterfaceMemberImplementation, declaration.IdentifierName, interfaceMember.ComponentName, interfaceMember.IdentifierName);

            var confirm = _messageBox.Show(message, RubberduckUI.ReorderParamsDialog_TitleText, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            return confirm == DialogResult.No ? null : interfaceMember;
        }

        private Declaration GetGetter()
        {
            if (TargetDeclaration == null) { return null; }

            if (TargetDeclaration.DeclarationType != DeclarationType.PropertyLet &&
                TargetDeclaration.DeclarationType != DeclarationType.PropertySet)
            {
                return TargetDeclaration;
            }

            var getter = Declarations.FirstOrDefault(item => item.Scope == TargetDeclaration.Scope &&
                                          item.IdentifierName == TargetDeclaration.IdentifierName &&
                                          item.DeclarationType == DeclarationType.PropertyGet);

            return getter ?? TargetDeclaration;
        }
    }
}
