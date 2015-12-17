using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.UI.IdentifierReferences;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Navigation
{
    public class NavigateAllImplementations : IDeclarationNavigator
    {
        private readonly VBE _vbe;
        private readonly AddIn _addIn;
        private readonly ICodePaneWrapperFactory _wrapperFactory;
        private readonly IMessageBox _messageBox;

        public NavigateAllImplementations(VBE vbe, AddIn addIn, ICodePaneWrapperFactory wrapperFactory, IMessageBox messageBox)
        {
            _vbe = vbe;
            _addIn = addIn;
            _wrapperFactory = wrapperFactory;
            _messageBox = messageBox;
        }

        public void Find()
        {
            //var codePane = _wrapperFactory.Create(_vbe.ActiveCodePane);
            //var selection = new QualifiedSelection(new QualifiedModuleName(codePane.CodeModule.Parent), codePane.Selection);
            //var progress = new ParsingProgressPresenter();
            //var parseResult = progress.Parse(_parser, _vbe.ActiveVBProject);

            //var implementsStatement = parseResult.Declarations.FindInterfaces()
            //    .SelectMany(i => i.References.Where(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext))
            //    .SingleOrDefault(r => r.QualifiedModuleName == selection.QualifiedName && r.Selection.Contains(selection.Selection));

            //if (implementsStatement != null)
            //{
            //    Find(implementsStatement.Declaration, parseResult);
            //}

            //var member = parseResult.Declarations.FindInterfaceImplementationMembers()
            //    .SingleOrDefault(m => m.Project == selection.QualifiedName.Project
            //                          && m.ComponentName == selection.QualifiedName.ComponentName
            //                          && m.Selection.Contains(selection.Selection)) ??
            //             parseResult.Declarations.FindInterfaceMembers()
            //                              .SingleOrDefault(m => m.Project == selection.QualifiedName.Project
            //                                                    && m.ComponentName == selection.QualifiedName.ComponentName
            //                                                    && m.Selection.Contains(selection.Selection));

            //if (member == null)
            //{
            //    return;
            //}

            //Find(member, parseResult);
        }

        public void Find(Declaration target)
        {
            //var progress = new ParsingProgressPresenter();
            //var parseResult = progress.Parse(_parser, _vbe.ActiveVBProject);
            //Find(target, parseResult);
        }

        private void Find(Declaration target, IRubberduckParserState parseResult)
        {
            var items = parseResult.AllDeclarations;
            string name;
            var implementations = (target.DeclarationType == DeclarationType.Class
                ? FindAllImplementationsOfClass(target, items, out name)
                : FindAllImplementationsOfMember(target, items, out name)) ?? new List<Declaration>();

            var declarations = implementations as IList<Declaration> ?? implementations.ToList();
            var implementationsCount = declarations.Count();

            if (implementationsCount == 1)
            {
                // if there's only 1 implementation, just jump to it:
                ImplementationsListDockablePresenter.OnNavigateImplementation(_vbe, declarations.First());
            }
            else if (implementationsCount > 1)
            {
                // if there's more than one implementation, show the dockable navigation window:
                try
                {
                    ShowImplementationsToolwindow(declarations, name);
                }
                catch (COMException)
                {
                    // the exception is related to the docked control host instance,
                    // trying again will work (I know, that's bad bad bad code)
                    ShowImplementationsToolwindow(declarations, name);
                }
            }
            else
            {
                var message = string.Format(RubberduckUI.AllImplementations_NoneFound, name);
                var caption = string.Format(RubberduckUI.AllImplementations_Caption, name);
                _messageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private IEnumerable<Declaration> FindAllImplementationsOfClass(Declaration target, IEnumerable<Declaration> declarations, out string name)
        {
            if (target.DeclarationType != DeclarationType.Class)
            {
                name = string.Empty;
                return null;
            }

            var identifiers = declarations as IList<Declaration> ?? declarations.ToList();

            var result = target.References
                .Where(reference => reference.Context.Parent is VBAParser.ImplementsStmtContext)
                .SelectMany(reference => identifiers.Where(identifier => identifier.IdentifierName == reference.QualifiedModuleName.ComponentName))
                .ToList();

            name = target.ComponentName;
            return result;
        }

        private IEnumerable<Declaration> FindAllImplementationsOfMember(Declaration target, IEnumerable<Declaration> declarations, out string name)
        {
            if (!target.DeclarationType.HasFlag(DeclarationType.Member))
            {
                name = string.Empty;
                return null;
            }

            var items = declarations as IList<Declaration> ?? declarations.ToList();

            var isInterface = items.FindInterfaces()
                .Select(i => i.QualifiedName.QualifiedModuleName.ToString())
                .Contains(target.QualifiedName.QualifiedModuleName.ToString());

            if (isInterface)
            {
                name = target.ComponentName + "." + target.IdentifierName;
                return items.FindInterfaceImplementationMembers(target.IdentifierName)
                       .Where(item => item.IdentifierName == target.ComponentName + "_" + target.IdentifierName);
            }

            var member = items.FindInterfaceMember(target);
            name = member.ComponentName + "." + member.IdentifierName;
            return items.FindInterfaceImplementationMembers(member.IdentifierName)
                   .Where(item => item.IdentifierName == member.ComponentName + "_" + member.IdentifierName);
        }

        private void ShowImplementationsToolwindow(IEnumerable<Declaration> implementations, string name)
        {
            // throws a COMException if toolwindow was already closed
            var window = new SimpleListControl(string.Format(RubberduckUI.AllImplementations_Caption, name));
            var presenter = new ImplementationsListDockablePresenter(_vbe, _addIn, window, implementations, _wrapperFactory);
            presenter.Show();
        }
    }
}