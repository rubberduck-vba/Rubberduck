using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Rename
{
    public interface IRename
    {
        void Rename(Declaration target, string newName);
        string ErrorMessage { get; }
        bool RequestParseAfterRename { get; }
    }

    public abstract class RenameRefactorBase : IRename
    {
        private readonly List<QualifiedModuleName> _modulesToRewrite;

        protected RenameRefactorBase()
        {
            _modulesToRewrite = new List<QualifiedModuleName>();
        }

        protected RenameRefactorBase(RubberduckParserState state)
        {
            State = state;
            _modulesToRewrite = new List<QualifiedModuleName>();
        }

        public RubberduckParserState State { get; }

        public void Rewrite()
        {
            foreach (var module in _modulesToRewrite.Distinct())
            {
                State.GetRewriter(module).Rewrite();
            }
        }

        public abstract void Rename(Declaration renameDeclaration, string newName);

        public abstract string ErrorMessage { get; }

        public virtual bool RequestParseAfterRename => true;

        public void RenameUsages(Declaration target, string newName)
        {
            var modules = target.References.GroupBy(r => r.QualifiedModuleName);
            foreach (var grouping in modules)
            {
                _modulesToRewrite.Add(grouping.Key);
                var rewriter = State.GetRewriter(grouping.Key);
                foreach (var reference in grouping)
                {
                    rewriter.Replace(reference.Context, newName);
                }
            }
        }

        public void RenameDeclaration(Declaration target, string newName)
        {
            _modulesToRewrite.Add(target.QualifiedName.QualifiedModuleName);
            var rewriter = State.GetRewriter(target.QualifiedName.QualifiedModuleName);

            if (target.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var members = State.DeclarationFinder.MatchName(target.IdentifierName)
                    .Where(item => item.ProjectId == target.ProjectId
                        && item.ComponentName == target.ComponentName
                        && item.DeclarationType.HasFlag(DeclarationType.Property));

                foreach (var member in members)
                {
                    var context = member.Context as IIdentifierContext;
                    if (null != context)
                    {
                        rewriter.Replace(context.IdentifierTokens, newName);
                    }
                }
            }
            else
            {
                var context = target.Context as IIdentifierContext;
                if (null != context)
                {
                    rewriter.Replace(context.IdentifierTokens, newName);
                }
            }
        }
    }
}
