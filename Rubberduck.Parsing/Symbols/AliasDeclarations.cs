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
            var conversionModule = _state.AllDeclarations.SingleOrDefault(
                    item => item.IdentifierName == "Conversion" && item.Scope == "VBE7.DLL;VBA.Conversion");

            var fileSystemModule = _state.AllDeclarations.SingleOrDefault(
                    item => item.IdentifierName == "FileSystem" && item.Scope == "VBE7.DLL;VBA.FileSystem");

            var interactionModule = _state.AllDeclarations.SingleOrDefault(
                    item => item.IdentifierName == "Interaction" && item.Scope == "VBE7.DLL;VBA.Interaction");

            var stringsModule = _state.AllDeclarations.SingleOrDefault(
                    item => item.IdentifierName == "Interaction" && item.Scope == "VBE7.DLL;VBA.Interaction");

            // all these modules are all part of the same project--only need to check one
            if (conversionModule == null)
            {
                return new List<Declaration>();
            }

            var functions = _state.AllDeclarations.Where(s => s.DeclarationType == DeclarationType.Function &&
                                                              s.Scope.StartsWith("VBE") &&
                                                              Tokens.Any(token => s.IdentifierName == "_B_var_" + token))
                                                  .ToList();

            var errorFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Conversion"), "Error"),
                    conversionModule,
                    conversionModule,
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

            var hexFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Conversion"), "Hex"),
                    conversionModule,
                    conversionModule,
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

            var octFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Conversion"), "Oct"),
                    conversionModule,
                    conversionModule,
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

            var strFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Conversion"), "Str"),
                    conversionModule,
                    conversionModule,
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

            var curDirFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "FileSystem"), "CurDir"),
                    fileSystemModule,
                    fileSystemModule,
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

            var commandFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Interaction"), "Command"),
                    interactionModule,
                    interactionModule,
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

            var environFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Interaction"), "Environ"),
                    interactionModule,
                    interactionModule,
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

            var chrFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Chr"),
                    stringsModule,
                    stringsModule,
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

            var chrwFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "ChrW"),
                    stringsModule,
                    stringsModule,
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

            var formatFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Format"),
                    stringsModule,
                    stringsModule,
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

            var lcaseFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "LCase"),
                    stringsModule,
                    stringsModule,
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

            var leftFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Left"),
                    stringsModule,
                    stringsModule,
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

            var leftbFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "LeftB"),
                    stringsModule,
                    stringsModule,
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

            var ltrimFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "LTrim"),
                    stringsModule,
                    stringsModule,
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

            var midFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Mid"),
                    stringsModule,
                    stringsModule,
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

            var midbFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "MidB"),
                    stringsModule,
                    stringsModule,
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

            var trimFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Trim"),
                    stringsModule,
                    stringsModule,
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

            var rightFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "Right"),
                    stringsModule,
                    stringsModule,
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

            var rightbFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "RightB"),
                    stringsModule,
                    stringsModule,
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

            var rtrimFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "RTrim"),
                    stringsModule,
                    stringsModule,
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

            var ucaseFunction = new FunctionDeclaration(
                    new QualifiedMemberName(
                        new QualifiedModuleName("VBA", "C:\\Program Files\\Common Files\\Microsoft Shared\\VBA\\VBA7.1\\VBE7.DLL", "Strings"), "UCase"),
                    stringsModule,
                    stringsModule,
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

            var functionAliases = new List<Declaration> {
                errorFunction,
                hexFunction,
                octFunction,
                strFunction,
                curDirFunction,
                commandFunction,
                environFunction,
                chrFunction,
                chrwFunction,
                formatFunction,
                lcaseFunction,
                leftFunction,
                leftbFunction,
                ltrimFunction,
                midFunction,
                midbFunction,
                trimFunction,
                rightFunction,
                rightbFunction,
                rtrimFunction,
                ucaseFunction
            };

            // ReSharper disable once PossibleInvalidCastExceptionInForeachLoop
            foreach (FunctionDeclaration alias in functionAliases)
            {
                foreach (var parameter in ((FunctionDeclaration)functions.Single(s => s.IdentifierName == "_B_var_" + alias.IdentifierName)).Parameters)
                {
                    alias.AddParameter(parameter);
                }
            }

            return functionAliases;
        }
    }
}