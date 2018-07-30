using Antlr4.Runtime;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Rewriter
{
    public class CodePaneRewriter : ModuleRewriterBase
    {
        public CodePaneRewriter(QualifiedModuleName module, ITokenStream tokenStream, IProjectsProvider projectsProvider)
            : base(module, tokenStream, projectsProvider)
        {}

        public override bool IsDirty
        {
            get
            {
                using (var codeModule = CodeModule())
                {
                    return codeModule == null || codeModule.Content() != Rewriter.GetText();
                }
            }
        }

        public override void Rewrite()
        {
            if (!IsDirty)
            {
                return;
            }

            using (var codeModule = CodeModule())
            {
                codeModule.Clear();
                var newContent = Rewriter.GetText();
                codeModule.InsertLines(1, newContent);
            }
        }
    }
}
