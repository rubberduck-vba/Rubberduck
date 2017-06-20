using System.IO;
using Antlr4.Runtime;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Rewriter
{
    /// <summary>
    /// A module rewriter that works off the MemberAttributes token stream and exports, overwrites and re-imports its module on <see cref="Rewrite"/>.
    /// This rewriter works off a token stream obtained from the AttributeParser, well before the code pane parse tree is acquired.
    /// </summary>
    /// <remarks>
    /// <ul>
    /// <li>DO NOT use this rewriter with any pending (not yet re-parsed) changes (e.g. refactorings, quick-fixes), or these changes will be lost.</li>
    /// <li>DO NOT use this rewriter to change any token that the VBE renders, or line number positions will be off.</li>
    /// <li>DO use this rewriter to add/remove hidden <c>Attribute</c> instructions to/from a module.</li>
    /// </ul>
    /// </remarks>
    public class MemberAttributesRewriter : ModuleRewriter
    {
        private readonly IModuleExporter _exporter;

        public MemberAttributesRewriter(IModuleExporter exporter, ICodeModule module, TokenStreamRewriter rewriter)
            : base(module, rewriter)
        {
            _exporter = exporter;
        }

        public override void Rewrite()
        {
            if(!IsDirty) { return; }

            var component = Module.Parent;
            if (component.Type == ComponentType.Document)
            {
                // can't re-import a document module
                return;
            }

            var file = _exporter.Export(component);
            var content = Rewriter.GetText();
            File.WriteAllText(file, content);

            var components = component.Collection;
            components.Remove(component);
            components.ImportSourceFile(file);
        }
    }
}