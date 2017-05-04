using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols.DeclarationLoaders
{
    public class AliasDeclarations : ICustomDeclarationLoader
    {
        private readonly RubberduckParserState _state;

        private Declaration _conversionModule;
        private Declaration _fileSystemModule;
        private Declaration _interactionModule;
        private Declaration _stringsModule;
        private Declaration _dateTimeModule;
        private Declaration _hiddenModule;

        public AliasDeclarations(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyList<Declaration> Load()
        {
            return AddAliasDeclarations();
        }

        private static readonly string[] Tokens =
        {
            Grammar.Tokens.Error,
            Grammar.Tokens.Hex,
            Grammar.Tokens.Oct,
            Grammar.Tokens.Str,
            Grammar.Tokens.StrConv,
            Grammar.Tokens.CurDir,
            Grammar.Tokens.Command,
            Grammar.Tokens.Environ,
            Grammar.Tokens.Chr,
            Grammar.Tokens.ChrW,
            Grammar.Tokens.Format,
            Grammar.Tokens.LCase,
            Grammar.Tokens.Left,
            Grammar.Tokens.LeftB,
            Grammar.Tokens.LTrim,
            Grammar.Tokens.Mid,
            Grammar.Tokens.MidB,
            Grammar.Tokens.Trim,
            Grammar.Tokens.Right,
            Grammar.Tokens.RightB,
            Grammar.Tokens.RTrim,
            Grammar.Tokens.String,
            Grammar.Tokens.UCase,
            Grammar.Tokens.Date,
            Grammar.Tokens.Time,
            Grammar.Tokens.Input,
            Grammar.Tokens.InputB
        };

        private IReadOnlyList<Declaration> AddAliasDeclarations()
        {
            var finder = _state.DeclarationFinder;;

            UpdateAliasFunctionModulesFromReferencedProjects(finder);

            if (NoReferenceToProjectContainingTheFunctionAliases() || WeHaveAlreadyLoadedTheDeclarationsBefore(finder))
            {
                return new List<Declaration>();
            }

            var possiblyAliasedFunctions = ReferencedBuiltInFunctionsThatMightHaveAnAlias(_state);
            var functionAliases = FunctionAliasesWithoutParameters();
            AddParametersToAliasesFromReferencedFunctions(functionAliases, possiblyAliasedFunctions);

            return functionAliases.Concat<Declaration>(PropertyGetDeclarations()).ToList();
        }

        private void UpdateAliasFunctionModulesFromReferencedProjects(DeclarationFinder finder)
        {
            var vba = finder.FindProject("VBA");
            if (vba == null)
            {
                // If the VBA project is null, we haven't loaded any COM references;
                // we're in a unit test and the mock project didn't setup any references.
                return;
            }

            _conversionModule = finder.FindStdModule("Conversion", vba, true);
            _fileSystemModule = finder.FindStdModule("FileSystem", vba, true);
            _interactionModule = finder.FindStdModule("Interaction", vba, true);
            _stringsModule = finder.FindStdModule("Strings", vba, true);
            _dateTimeModule = finder.FindStdModule("DateTime", vba, true);
            _hiddenModule = finder.FindStdModule("_HiddenModule", vba, true);
        }

        private bool NoReferenceToProjectContainingTheFunctionAliases()
        {
            return _conversionModule == null;
            // All the modules containing function aliases are part of the same project. --> Only need to check one.
        }

        private bool WeHaveAlreadyLoadedTheDeclarationsBefore(DeclarationFinder finder)
        {
            return ThereIsAnErrorFunctionDeclaration(finder);
        }

        private bool ThereIsAnErrorFunctionDeclaration(DeclarationFinder finder)
        {
            var errorFunction = ErrorFunction();
            return finder.MatchName(errorFunction.IdentifierName)
                            .Any(declaration => declaration.Equals(errorFunction));
        }

        private List<Declaration> ReferencedBuiltInFunctionsThatMightHaveAnAlias(RubberduckParserState state)
        {
            var functions = state.AllDeclarations.Where(s => s.DeclarationType == DeclarationType.Function
                                                             && s.Scope.StartsWith("VBE")
                                                             &&
                                                             Tokens.Any(token => s.IdentifierName == "_B_var_" + token));
            return functions.ToList();
        }

        private List<PropertyGetDeclaration> PropertyGetDeclarations()
        {
            return new List<PropertyGetDeclaration>
            {
                DatePropertyGet(),
                TimePropertyGet(),
            };
        }

        private PropertyGetDeclaration DatePropertyGet()
        {
            return new PropertyGetDeclaration(
                new QualifiedMemberName(_dateTimeModule.QualifiedName.QualifiedModuleName, "Date"),
                _dateTimeModule,
                _dateTimeModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private PropertyGetDeclaration TimePropertyGet()
        {
            return new PropertyGetDeclaration(
                new QualifiedMemberName(_dateTimeModule.QualifiedName.QualifiedModuleName, "Time"),
                _dateTimeModule,
                _dateTimeModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private List<FunctionDeclaration> FunctionAliasesWithoutParameters()
        {
            return new List<FunctionDeclaration>
            {
                ErrorFunction(),
                HexFunction(),
                OctFunction(),
                StrFunction(),
                StrConvFunction(),
                CurDirFunction(),
                CommandFunction(),
                EnvironFunction(),
                ChrFunction(),
                ChrwFunction(),
                FormatFunction(),
                LCaseFunction(),
                LeftFunction(),
                LeftBFunction(),
                LTrimFunction(),
                MidFunction(),
                MidBFunction(),
                TrimFunction(),
                RightFunction(),
                RightBFunction(),
                RTrimFunction(),
                StringFunction(),
                UCaseFunction(),
                InputFunction(),
                InputBFunction(),
            };
        }

        private FunctionDeclaration ErrorFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_conversionModule.QualifiedName.QualifiedModuleName, "Error"),
                _conversionModule,
                _conversionModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration HexFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_conversionModule.QualifiedName.QualifiedModuleName, "Hex"),
                _conversionModule,
                _conversionModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration OctFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_conversionModule.QualifiedName.QualifiedModuleName, "Oct"),
                _conversionModule,
                _conversionModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration StrFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_conversionModule.QualifiedName.QualifiedModuleName, "Str"),
                _conversionModule,
                _conversionModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration StrConvFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "StrConv"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration CurDirFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_fileSystemModule.QualifiedName.QualifiedModuleName, "CurDir"),
                _fileSystemModule,
                _fileSystemModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration CommandFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_interactionModule.QualifiedName.QualifiedModuleName, "Command"),
                _interactionModule,
                _interactionModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration EnvironFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_interactionModule.QualifiedName.QualifiedModuleName, "Environ"),
                _interactionModule,
                _interactionModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration ChrFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "Chr"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration ChrwFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "ChrW"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration FormatFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "Format"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration LCaseFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "LCase"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration LeftFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "Left"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration LeftBFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "LeftB"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration LTrimFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "LTrim"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration MidFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "Mid"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration MidBFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "MidB"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration TrimFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "Trim"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration RightFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "Right"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration RightBFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "RightB"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration RTrimFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "RTrim"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration StringFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "String"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration UCaseFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_stringsModule.QualifiedName.QualifiedModuleName, "UCase"),
                _stringsModule,
                _stringsModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration InputFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_hiddenModule.QualifiedName.QualifiedModuleName, "Input"),
                _hiddenModule,
                _hiddenModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private FunctionDeclaration InputBFunction()
        {
            return new FunctionDeclaration(
                new QualifiedMemberName(_hiddenModule.QualifiedName.QualifiedModuleName, "InputB"),
                _hiddenModule,
                _hiddenModule,
                "Variant",
                null,
                string.Empty,
                Accessibility.Global,
                null,
                new Selection(),
                false,
                false,
                new List<IAnnotation>(),
                new Attributes());
        }

        private static void AddParametersToAliasesFromReferencedFunctions(List<FunctionDeclaration> functionAliases, List<Declaration> referencedFunctions)
        {
            foreach (var alias in functionAliases)
            {
                var function = referencedFunctions.OfType<FunctionDeclaration>()
                    .SingleOrDefault(s => s.IdentifierName == "_B_var_" + alias.IdentifierName);

                if (function == null) { continue; }
                foreach (var parameter in function.Parameters)
                {
                    alias.AddParameter(parameter);
                }
            }
        }
    }
}