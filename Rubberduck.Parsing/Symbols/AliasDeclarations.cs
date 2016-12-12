using System.Collections.Generic;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Linq;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.Parsing.Symbols
{
    public class AliasDeclarations : ICustomDeclarationLoader
    {
        private readonly RubberduckParserState _state;

        private Declaration _conversionModule;
        private Declaration _fileSystemModule;
        private Declaration _interactionModule;
        private Declaration _stringsModule;


        public AliasDeclarations(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyList<Declaration> Load()
        {
            return AddAliasDeclarations();
        }

        private static readonly string[] Tokens = {
            Grammar.Tokens.Error,
            Grammar.Tokens.Hex,
            Grammar.Tokens.Oct,
            Grammar.Tokens.Str,
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
            Grammar.Tokens.UCase
        };

        private IReadOnlyList<Declaration> AddAliasDeclarations()
        {
            UpdateAliasFunctionModulesFromReferencedProjects(_state);
            
            if (NoReferenceToProjectContainingTheFunctionAliases())
            {
                return new List<Declaration>();
            }
            
            var possiblyAliasedFunctions = ReferencedBuiltInFunctionsThatMightHaveAnAlias(_state);
            var functionAliases = FunctionAliasesWithoutParameters();
            AddParametersToAliasesFromReferencedFunctions(functionAliases, possiblyAliasedFunctions);
            
            return functionAliases;
        }

            private void UpdateAliasFunctionModulesFromReferencedProjects(RubberduckParserState state)
            {
                _conversionModule = state.AllDeclarations.SingleOrDefault(
                        item => item.IdentifierName == "Conversion" && item.Scope == "VBE7.DLL;VBA.Conversion");

                _fileSystemModule = state.AllDeclarations.SingleOrDefault(
                        item => item.IdentifierName == "FileSystem" && item.Scope == "VBE7.DLL;VBA.FileSystem");

                _interactionModule = state.AllDeclarations.SingleOrDefault(
                        item => item.IdentifierName == "Interaction" && item.Scope == "VBE7.DLL;VBA.Interaction");

                _stringsModule = state.AllDeclarations.SingleOrDefault(
                        item => item.IdentifierName == "Strings" && item.Scope == "VBE7.DLL;VBA.Strings");
            }


            private bool NoReferenceToProjectContainingTheFunctionAliases()
            {
                return _conversionModule == null;   // All the modules containing function aliases are part of the same project. --> Only need to check one.
            }


            private List<Declaration> ReferencedBuiltInFunctionsThatMightHaveAnAlias(RubberduckParserState state)
            {
                var functions = state.AllDeclarations.Where(s => s.DeclarationType == DeclarationType.Function
                                                                  && s.Scope.StartsWith("VBE")
                                                                  && Tokens.Any(token => s.IdentifierName == "_B_var_" + token));
                return functions.ToList();
            }


            private List<FunctionDeclaration> FunctionAliasesWithoutParameters()
            {
                return new List<FunctionDeclaration> {
                    ErrorFunction(),
                    HexFunction(),
                    OctFunction(),
                    StrFunction(),
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
                    UCaseFunction()
                };
            }

                private FunctionDeclaration ErrorFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Conversion"), "Error"),
                            _conversionModule,
                            _conversionModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration HexFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Conversion"), "Hex"),
                            _conversionModule,
                            _conversionModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration OctFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Conversion"), "Oct"),
                            _conversionModule,
                            _conversionModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration StrFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Conversion"), "Str"),
                            _conversionModule,
                            _conversionModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration CurDirFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "FileSystem"), "CurDir"),
                            _fileSystemModule,
                            _fileSystemModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration CommandFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Interaction"), "Command"),
                            _interactionModule,
                            _interactionModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration EnvironFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Interaction"), "Environ"),
                            _interactionModule,
                            _interactionModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration ChrFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Chr"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration ChrwFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "ChrW"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration FormatFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Format"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration LCaseFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "LCase"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration LeftFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Left"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration LeftBFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "LeftB"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration LTrimFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "LTrim"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration MidFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Mid"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration MidBFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "MidB"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration TrimFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Trim"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration RightFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Right"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration RightBFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "RightB"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration RTrimFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "RTrim"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }

                private FunctionDeclaration UCaseFunction()
                {
                    return new FunctionDeclaration(
                            new QualifiedMemberName(
                                new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "UCase"),
                            _stringsModule,
                            _stringsModule,
                            "Variant",
                            null,
                            string.Empty,
                            Accessibility.Global,
                            null,
                            new Selection(),
                            false,
                            true,
                            new List<IAnnotation>(),
                            new Attributes());
                }


            private static void AddParametersToAliasesFromReferencedFunctions(List<FunctionDeclaration> functionAliases, List<Declaration> referencedFunctions)
            {
                // ReSharper disable once PossibleInvalidCastExceptionInForeachLoop
                foreach (var alias in functionAliases)
                {
                    foreach (var parameter in ((FunctionDeclaration)referencedFunctions.Single(s => s.IdentifierName == "_B_var_" + alias.IdentifierName)).Parameters)
                    {
                        alias.AddParameter(parameter);
                    }
                }
            }

    }
}