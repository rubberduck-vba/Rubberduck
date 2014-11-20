using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class ProcedureNode : CodeBlockNode
    {
        public ProcedureNode(Instruction instruction, string scope, Match match, string keyword, IEnumerable<SyntaxTreeNode> nodes)
            : base(instruction, scope, match, new[] {ReservedKeywords.End + " " + keyword}, null, nodes)
        {
            _identifier = CreateIdentifier(scope, match);
            _parameters = CreateParameters(scope + '.' +  _identifier.Name, match).ToList();
        }

        private Identifier CreateIdentifier(string scope, Match match)
        {
            var name = match.Groups["identifier"].Captures[0].Value;

            var kind = match.Groups["kind"].Value;
            var hasReturnType = kind == ReservedKeywords.Function ||
                                (kind == ReservedKeywords.Property && kind.EndsWith(ReservedKeywords.Get));

            var specifiedType = match.Groups["reference"];
            var returnType = hasReturnType
                ? specifiedType.Success ? specifiedType.Value : ReservedKeywords.Variant
                : null;

            return new Identifier(scope, name, returnType);
        }

        private IEnumerable<ParameterNode> CreateParameters(string scope, Match match)
        {
            var parameters = match.Groups["parameters"].Value.Split(',');
            var pattern = VBAGrammar.ParameterSyntax;
            foreach (var parameter in parameters)
            {
                var subMatch = Regex.Match(parameter, pattern);
                var startColumn = Instruction.Value.IndexOf('(') + 1 + subMatch.Index;
                var endColumn = startColumn + subMatch.Length;
                var instruction = new Instruction(Instruction.Line, startColumn, endColumn, subMatch.Value.Replace(",", string.Empty));

                yield return new ParameterNode(instruction, scope, subMatch);
            }
        }

        private readonly Identifier _identifier;
        public Identifier Identifier { get { return _identifier; } }

        private readonly IEnumerable<ParameterNode> _parameters;
        public IEnumerable<ParameterNode> Parameters { get { return _parameters; } }
    }
}
