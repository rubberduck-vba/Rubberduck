using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    public interface IRenameRefactoringHandler
    {
        void Rename();
        string ErrorMessage { get; }
    }

    public class RenameHandlerBase : IRenameRefactoringHandler
    {
        private readonly RenameModel _model;
        private readonly IMessageBox _messageBox;
        private List<QualifiedModuleName> _updatedQMNs;

        public RenameHandlerBase(RenameModel model, IMessageBox messageBox)
        {
            _model = model;
            _messageBox = messageBox;
            _updatedQMNs = new List<QualifiedModuleName>();
        }

        public RenameModel Model { get { return _model; } }

        public IMessageBox MessageBox { get { return _messageBox; } }

        public void Rewrite()
        {
            _updatedQMNs.Distinct().ToList()
                .ForEach(qmn => Model.State.GetRewriter(qmn).Rewrite());
        }

        public virtual void Rename() { }
        public virtual string ErrorMessage { get; }

        public void RenameUsages(Declaration target)
        {
            RenameUsages(target, _model.NewName);
        }

        public void RenameUsages(Declaration target, string newName)
        {
            var qualifiedModuleNames = new List<QualifiedModuleName>();
            var modules = target.References.GroupBy(r => r.QualifiedModuleName);
            foreach (var grouping in modules)
            {
                qualifiedModuleNames.Add(grouping.Key);
                var rewriter = _model.State.GetRewriter(grouping.Key);
                var module = grouping.Key.Component.CodeModule;
                foreach (var line in grouping.GroupBy(reference => reference.Selection.StartLine))
                {
                    var lastSelection = Selection.Empty;
                    foreach (var reference in line.OrderByDescending(l => l.Selection.StartColumn))
                    {
                        if (reference.Selection == lastSelection)
                        {
                            continue;
                        }
                        var newContent = reference.Context.GetText().Replace(reference.IdentifierName, newName);
                        rewriter.Replace(reference.Context, newContent);
                        lastSelection = reference.Selection;
                    }
                }
            }
            _updatedQMNs.AddRange(qualifiedModuleNames);
        }

        public void RenameDeclaration(Declaration target)
        {
            RenameDeclaration(target, _model.NewName);
        }

        public void RenameDeclaration(Declaration target, string newName)
        {
            var qualifiedModuleName = target.QualifiedName.QualifiedModuleName;
            var component = qualifiedModuleName.Component;
            var rewriter = _model.State.GetRewriter(target);
            var module = component.CodeModule;
            if (!target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var newContent = target.Context.GetText().Replace(target.IdentifierName, newName);
                rewriter.Replace(target.Context, newContent);
            }
            else
            {
                var members = _model.Declarations.Named(target.IdentifierName)
                    .Where(item => item.ProjectId == target.ProjectId
                        && item.ComponentName == target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property));

                foreach (var member in members)
                {
                    var newContent = member.Context.GetText().Replace(member.IdentifierName, newName);
                    rewriter.Replace(member.Context, newContent);
                }
            }
            _updatedQMNs.Add(qualifiedModuleName);
        }
    }
}
