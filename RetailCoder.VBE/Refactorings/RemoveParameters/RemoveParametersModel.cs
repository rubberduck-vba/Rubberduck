using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersModel
    {
        private readonly VBProjectParseResult _parseResult;
        public VBProjectParseResult ParseResult { get { return _parseResult; } }

        private readonly Declarations _declarations;
        public Declarations Declarations { get { return _declarations; } }

        public Declaration TargetDeclaration { get; set; }
        public List<Parameter> Parameters { get; set; }

        public RemoveParametersModel(VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            _parseResult = parseResult;
            _declarations = parseResult.Declarations;

            TargetDeclaration = FindTarget(selection, ValidDeclarationTypes);

            Parameters = new List<Parameter>();
            LoadParameters();
        }

        public void LoadParameters()
        {
            Parameters.Clear();

            var index = 0;
            Parameters = GetParameters(TargetDeclaration).Select(arg => new Parameter(arg, index++)).ToList();
        }

        private IEnumerable<Declaration> GetParameters(Declaration method)
        {
            return Declarations.Items
                              .Where(d => d.DeclarationType == DeclarationType.Parameter
                                       && d.ComponentName == method.ComponentName
                                       && d.Project.Equals(method.Project)
                                       && method.Context.Start.Line <= d.Selection.StartLine
                                       && method.Context.Stop.Line >= d.Selection.EndLine
                                       && !(method.Context.Start.Column > d.Selection.StartColumn && method.Context.Start.Line == d.Selection.StartLine)
                                       && !(method.Context.Stop.Column < d.Selection.EndColumn && method.Context.Stop.Line == d.Selection.EndLine))
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

        private Declaration FindTarget(QualifiedSelection selection, DeclarationType[] validDeclarationTypes)
        {
            var target = Declarations.Items
                .Where(item => !item.IsBuiltIn)
                .FirstOrDefault(item => IsSelectedDeclaration(selection, item)
                                        || IsSelectedReference(selection, item));

            if (target != null && validDeclarationTypes.Contains(target.DeclarationType))
            {
                return target;
            }

            target = null;

            var targets = Declarations.Items
                .Where(item => !item.IsBuiltIn
                               && item.ComponentName == selection.QualifiedName.ComponentName
                               && validDeclarationTypes.Contains(item.DeclarationType));

            var currentSelection = new Selection(0, 0, int.MaxValue, int.MaxValue);

            foreach (var declaration in targets)
            {
                var declarationSelection = new Selection(declaration.Context.Start.Line,
                    declaration.Context.Start.Column,
                    declaration.Context.Stop.Line,
                    declaration.Context.Stop.Column);

                if (currentSelection.Contains(declarationSelection) &&
                    declarationSelection.Contains(selection.Selection))
                {
                    target = declaration;
                    currentSelection = declarationSelection;
                }

                foreach (var reference in declaration.References)
                {
                    var proc = (dynamic) reference.Context.Parent;
                    VBAParser.ArgsCallContext paramList;

                    // This is to prevent throws when this statement fails:
                    // (VBAParser.ArgsCallContext)proc.argsCall();
                    try
                    {
                        paramList = (VBAParser.ArgsCallContext) proc.argsCall();
                    }
                    catch
                    {
                        continue;
                    }

                    if (paramList == null)
                    {
                        continue;
                    }

                    var referenceSelection = new Selection(paramList.Start.Line,
                        paramList.Start.Column,
                        paramList.Stop.Line,
                        paramList.Stop.Column + paramList.Stop.Text.Length + 1);

                    if (currentSelection.Contains(declarationSelection) &&
                        referenceSelection.Contains(selection.Selection))
                    {
                        target = reference.Declaration;
                        currentSelection = referenceSelection;
                    }
                }
            }
            return target;
        }

        private bool IsSelectedReference(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.References.Any(r =>
                r.QualifiedModuleName == selection.QualifiedName &&
                r.Selection.ContainsFirstCharacter(selection.Selection));
        }

        private bool IsSelectedDeclaration(QualifiedSelection selection, Declaration declaration)
        {
            return declaration.QualifiedName.QualifiedModuleName == selection.QualifiedName
                   && (declaration.Selection.ContainsFirstCharacter(selection.Selection));
        }
    }
}
