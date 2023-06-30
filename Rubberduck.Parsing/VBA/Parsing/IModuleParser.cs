using System.Collections.Generic;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.Parsing
{
    public readonly struct ModuleParseResults
    {
        public ModuleParseResults(IParseTree codePaneParseTree,
            IParseTree attributesParseTree,
            IEnumerable<CommentNode> comments,
            IEnumerable<IParseTreeAnnotation> annotations,
            LogicalLineStore logicalLines,
            IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> attributes,
            IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext>
                membersAllowingAttributes,
            ITokenStream codePaneTokenStream,
            ITokenStream attributesTokenStream)
        {
            CodePaneParseTree = codePaneParseTree;
            AttributesParseTree = attributesParseTree;
            Comments = comments;
            Annotations = annotations;
            Attributes = attributes;
            MembersAllowingAttributes = membersAllowingAttributes;
            CodePaneTokenStream = codePaneTokenStream;
            AttributesTokenStream = attributesTokenStream;
            LogicalLines = logicalLines;
        }

        public IParseTree CodePaneParseTree { get; }
        public IParseTree AttributesParseTree { get; }
        public IEnumerable<CommentNode> Comments { get; }
        public IEnumerable<IParseTreeAnnotation> Annotations { get; }
        public LogicalLineStore LogicalLines { get; }
        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), Attributes> Attributes { get; }
        public IDictionary<(string scopeIdentifier, DeclarationType scopeType), ParserRuleContext> MembersAllowingAttributes { get; }
        public ITokenStream CodePaneTokenStream { get; }
        public ITokenStream AttributesTokenStream { get; }
    }

    public interface IModuleParser
    {
        ModuleParseResults Parse(QualifiedModuleName module, CancellationToken cancellationToken, TokenStreamRewriter rewriter = null);
    }
}
