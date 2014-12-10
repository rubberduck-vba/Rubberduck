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

            var kind = match.Groups["kind"].Value;
            _hasReturnType = kind == ReservedKeywords.Function ||
                             (kind == ReservedKeywords.Property && kind.EndsWith(ReservedKeywords.Get));

            _specifiedReturnType = match.Groups["reference"].Value;
        }

        private readonly IEnumerable<ParameterNode> _parameters;
        public IEnumerable<ParameterNode> Parameters { get { return _parameters; } }

        private Identifier CreateIdentifier(string scope, Match match)
        {
            var name = match.Groups["identifier"].Captures[0].Value;

            var returnType = HasReturnType
                ? string.IsNullOrEmpty(SpecifiedReturnType) ? ReservedKeywords.Variant : SpecifiedReturnType
                : null;

            return new Identifier(scope, name, returnType);
        }

        private IEnumerable<ParameterNode> CreateParameters(string scope, Match match)
        {
            var parameters = match.Groups["parameter"].Captures.Cast<Capture>();
            var pattern = VBAGrammar.ParameterSyntax;
            var caret = 0;
            foreach (var parameter in parameters.Where(p => !string.IsNullOrEmpty(p.Value)).OrderBy(p => p.Index))
            {
                // todo: stop assuming the instruction sits on a single line of code _
                var subMatch = Regex.Match(parameter.Value, pattern);
                var startColumn = caret + Instruction.Value.IndexOf('(') + 2; // 1 to move after '(', 1 because we want 1-based column number
                var endColumn = startColumn + subMatch.Value.Length;
                caret = endColumn + ", ".Length - startColumn;
                var instruction = new Instruction(Instruction.Line, startColumn, endColumn, subMatch.Value.Replace(",", string.Empty));

                yield return new ParameterNode(instruction, scope, subMatch);
            }
        }

        private bool _hasReturnType;
        private string _specifiedReturnType;

        private readonly Identifier _identifier;
        public Identifier Identifier { get { return _identifier; } }

        public string Accessibility { get { return GetAccessibility(); } }

        private string GetAccessibility()
        {
            var keywords = new[] {ReservedKeywords.Public, ReservedKeywords.Private, ReservedKeywords.Friend};
            var value = RegexMatch.Groups["accessibility"].Value;

            return (keywords.Contains(value)) ? value : ReservedKeywords.Public; // VBA procs are public by default
        }

        public ProcedureKind Kind
        {
            get
            {
                var kind = RegexMatch.Groups["kind"].Value;
                return kind.StartsWith(ReservedKeywords.Property)
                    ? kind.EndsWith(ReservedKeywords.Get)
                        ? ProcedureKind.PropertyGet
                        : kind.EndsWith(ReservedKeywords.Let)
                            ? ProcedureKind.PropertyLet
                            : ProcedureKind.PropertySet
                    : kind == ReservedKeywords.Sub ? ProcedureKind.Sub : ProcedureKind.Function;
            }
        }

        public bool HasReturnType
        {
            get { return _hasReturnType; }
        }

        public string SpecifiedReturnType
        {
            get { return _specifiedReturnType; }
        }
    }
}
