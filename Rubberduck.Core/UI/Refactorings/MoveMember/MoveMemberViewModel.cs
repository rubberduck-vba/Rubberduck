using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Rubberduck.UI.Refactorings.MoveMember
{
    public class MoveMemberViewModel : RefactoringViewModelBase<MoveMemberModel>
    {
        private const string _instructions = "Specify Members to Move and a Destination Module";

        public MoveMemberViewModel(MoveMemberModel model)
            : base(model) { }

        public string RefactorName => "Relocate a Members to a new or existing Procedural Module";
        public string Instructions => _instructions;

        private string ConflictHeaderFormat => "Moving {0} is not an executable move for the following reason(s):";
        private string ConflictNullMemberIdentifier => "the method";
        private string PreviewHeaderFormat => "Moving '{0}' will move the following element(s) to Module: {1}";
        private string NoPreviewHeaderFormat => "Moving '{0}' to '{1}' is an executable move";
        private string NoMoveableMembersFormat => "{0} has no moveable members";

        public Func<IEnumerable<string>> ConflictsRetriever { set; get; } 
            = () => new List<string>() {_instructions};

        public bool IsExecutableMove => !ConflictsRetriever().Any();

        //public Func<Declaration, IDeclarationFinderProvider, Declaration> DefaultCandidate { set; get; }

        private Declaration _memberToMove;
        public Declaration MemberToMove
        {
            get
            {
                return _memberToMove;
            }
            set
            {
                _memberToMove = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsExecutableMove));
                OnPropertyChanged(nameof(MoveCommentary));
            }
        }

        private Declaration _sourceModule;
        public Declaration SourceModule
        {
            get
            {
                return _sourceModule;
            }
            set
            {
                if (_sourceModule == value) { return; }

                _sourceModule = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(MoveCandidates));
                OnPropertyChanged(nameof(IsExecutableMove));
                OnPropertyChanged(nameof(MoveCommentary));
                OnPropertyChanged(nameof(DestinationModules));
                MemberToMove = null; // DefaultCandidate is null ? null : DefaultCandidate(value, Model.State);
            }
        }

        string _destinationModuleName;
        public string DestinationModuleName
        {
            get
            {
                return DestinationModule?.IdentifierName ?? _destinationModuleName ?? string.Empty;
            }
            set
            {
                _destinationModuleName = value;
                OnPropertyChanged(nameof(IsExecutableMove));
                OnPropertyChanged(nameof(MoveCommentary));
            }
        }

        public Declaration DestinationModule { set; get; }

        public IEnumerable<KeyValuePair<Declaration, string>> DestinationModules
            => Modules(DeclarationType.ProceduralModule)
            .Where(mod => mod.Key != SourceModule);

        public IEnumerable<KeyValuePair<Declaration, string>> SourceModules
            => Modules(DeclarationType.Module);

        private IEnumerable<KeyValuePair<Declaration, string>> Modules(Enum typeFlag)
        {
            string moduleDisplayName(Declaration mod)
            {
                if (mod.QualifiedModuleName.ComponentType == ComponentType.ClassModule)
                {
                    return $"{mod.IdentifierName} ({Localize("ClassModule")})";
                }
                if (mod.QualifiedModuleName.ComponentType == ComponentType.UserForm)
                {
                    //TODO: Add "form" to RubberduckUI.resx
                    return $"{mod.IdentifierName} (form)";
                }
                return $"{mod.IdentifierName} ({Localize("ProceduralModule")})";
            }

            return Model.DeclarationFinderProvider.DeclarationFinder.AllUserDeclarations
            .Where(ud => ud.DeclarationType.HasFlag(typeFlag))
            .OrderBy(ud => ud.IdentifierName)
            .Select(mod => new KeyValuePair<Declaration, string>(mod, moduleDisplayName(mod)));
        }

        public IEnumerable<KeyValuePair<Declaration,string>> MoveCandidates
        {
            get
            {
                if (SourceModule is null)
                {
                    return Enumerable.Empty<KeyValuePair<Declaration, string>>();
                }

                var allMembers = Model.DeclarationFinderProvider.DeclarationFinder.AllUserDeclarations.Where(member => member.ProjectId == SourceModule.ProjectId
                           && member.ComponentName == SourceModule.ComponentName);

                var members = allMembers.Where(m => m.DeclarationType.HasFlag(DeclarationType.Member))
                    .Select(d => new KeyValuePair<Declaration, string>(d, $"{MemberDisplaySignature(d)}"))
                    .OrderBy(kv => kv.Key.IdentifierName);

                var constants = allMembers.Where(m => m.IsConstant()
                    && !m.IsLocalConstant())
                    .Select(d => new KeyValuePair<Declaration, string>(d, $"{d.Accessibility.ToString()} {Tokens.Const} {d.IdentifierName} As {d.AsTypeName}"))
                    .OrderBy(kv => kv.Key.IdentifierName);

                var variables = allMembers.Where(m => m.IsVariable()
                    && !m.IsLocalVariable() )
                    .Select(d => new KeyValuePair<Declaration, string>(d, $"{d.Accessibility.ToString()} {d.IdentifierName} As {d.AsTypeName}"))
                    .OrderBy(kv => kv.Key.IdentifierName);

                var enumerations = allMembers.Where(m => m.DeclarationType.HasFlag(DeclarationType.Enumeration))
                    .Select(d => new KeyValuePair<Declaration, string>(d, $"{d.Accessibility.ToString()} {d.IdentifierName} As {ToLocalizedString(d.DeclarationType)}"))
                    .OrderBy(kv => kv.Key.IdentifierName);

                var userDefinedTypes = allMembers.Where(m => m.DeclarationType.HasFlag(DeclarationType.UserDefinedType))
                    .Select(d => new KeyValuePair<Declaration, string>(d, $"{d.Accessibility.ToString()} {d.IdentifierName} As {ToLocalizedString(d.DeclarationType)}"))
                    .OrderBy(kv => kv.Key.IdentifierName);

                return members.Concat(constants)
                                .Concat(variables)
                                .Concat(enumerations)
                                .Concat(userDefinedTypes);
            }
        }

        private string ToLocalizedString(DeclarationType type)
            => RubberduckUI.ResourceManager.GetString("DeclarationType_" + type, CultureInfo.CurrentUICulture);


        private string MemberDisplaySignature(Declaration member)
        {
            if (member is ModuleBodyElementDeclaration moduleBodyElementDeclaration)
            {
                if (moduleBodyElementDeclaration.Accessibility.Equals(Accessibility.Implicit))
                {
                    return ImprovedSignature(moduleBodyElementDeclaration, Accessibility.Public);
                }

                return ImprovedSignature(moduleBodyElementDeclaration);
            }
            return  member.IdentifierName;
        }

        public Func<string> Preview { set; get; }

        public string MoveCommentary
        {
            get
            {
                if (!MoveCandidates.Any())
                {
                    return string.Format(NoMoveableMembersFormat, SourceModule.IdentifierName);
                }

                var conflicts = ConflictsRetriever();

                if (!conflicts.Any())
                {
                    var movedContent = Preview();
                    var previewHeaderFormat = movedContent.Length > 0 ? PreviewHeaderFormat : NoPreviewHeaderFormat;
                    var header = string.Format(previewHeaderFormat, MemberToMove.IdentifierName, DestinationModuleName);
                    return $"{header}{Environment.NewLine}{movedContent}";
                }

                var commentary = string.Empty;
                var num = 1;
                foreach (var conflict in conflicts)
                {
                    commentary = $"{commentary}{Environment.NewLine}{num++}. {conflict}";
                }
                var memberDecriptor = MemberToMove is null ? ConflictNullMemberIdentifier : $"'{MemberToMove.IdentifierName}'";
                return $"{string.Format(ConflictHeaderFormat, $"{memberDecriptor}")}{Environment.NewLine}{Environment.NewLine}{commentary}";
            }
        }

        private string Localize(string type)
            => RubberduckUI.ResourceManager.GetString("DeclarationType_" + type, CultureInfo.CurrentUICulture);

        /// <summary>
        /// Returns a member's signature with an improved argument list.
        /// </summary>
        private static string ImprovedSignature(ModuleBodyElementDeclaration declaration)
            => ImprovedSignature(declaration, declaration.Accessibility);

        /// <summary>
        /// Returns a member's signature with specified accessibility and an improved argument list.
        /// </summary>
        private static string ImprovedSignature(ModuleBodyElementDeclaration declaration, Accessibility accessibility)
        {
            var memberType = string.Empty;
            switch (declaration.Context)
            {
                case VBAParser.SubStmtContext _:
                    memberType = Tokens.Sub;
                    break;
                case VBAParser.FunctionStmtContext _:
                    memberType = Tokens.Function;
                    break;
                case VBAParser.PropertyGetStmtContext _:
                    memberType = $"{Tokens.Property} {Tokens.Get}";
                    break;
                case VBAParser.PropertyLetStmtContext _:
                    memberType = $"{Tokens.Property} {Tokens.Let}";
                    break;
                case VBAParser.PropertySetStmtContext _:
                    memberType = $"{Tokens.Property} {Tokens.Set}";
                    break;
                default:
                    throw new ArgumentException();
            }

            var accessibilityToken = accessibility.Equals(Accessibility.Implicit) ? string.Empty : $"{accessibility.ToString()}";

            var signature = $"{memberType} {declaration.IdentifierName}({ImprovedParameterList(declaration)})";

            var fullSignature = declaration.AsTypeName == null ?
                $"{accessibilityToken} {signature}"
                : $"{accessibilityToken} {signature} As {declaration.AsTypeName}";

            return fullSignature;
        }


        //Modifies parameter list based on
        //considerations identified in https://github.com/rubberduck-vba/Rubberduck/issues/3486
        /// <summary>
        /// Returns a member's parameter list improved with explicit ByRef/ByVal parameter modifiers
        /// and Type identifiers.
        /// </summary>
        public static string ImprovedParameterList(ModuleBodyElementDeclaration declaration)
        {
            var memberParams = new List<string>();
            if (declaration is IParameterizedDeclaration memberWithParams)
            {
                var paramAttributeTuples = memberWithParams.Parameters.TakeWhile(p => p != memberWithParams.Parameters.Last())
                    .OrderBy(o => o.Selection.StartLine)
                    .ThenBy(t => t.Selection.StartColumn)
                    .Select(p => ExamineArgContext(p)).ToList();

                if (memberWithParams.Parameters.Any())
                {
                    var lastParamRequiresByValModifier = declaration.DeclarationType.HasFlag(DeclarationType.PropertyLet)
                        || declaration.DeclarationType.HasFlag(DeclarationType.PropertySet);
                    paramAttributeTuples.Add(ExamineArgContext(memberWithParams.Parameters.Last(), lastParamRequiresByValModifier));
                }

                memberParams = paramAttributeTuples.Select(ps =>
                    ParameterTupleToString(
                        ps.Modifier,
                        ps.Name,
                        ps.Type)).ToList();
            }

            return string.Join(", ", memberParams);

            //local methods below
            string ParameterTupleToString(string paramModifier, string paramName, string paramType)
                => paramModifier.Length > 0 ? $"{paramModifier} {paramName} As {paramType}" : $"{paramName} As {paramType}";

            (string Modifier, string Name, string Type) ExamineArgContext(ParameterDeclaration param, bool requiresByVal = false)
            {
                var arg = param.Context as VBAParser.ArgContext;
                var modifier = arg.BYVAL()?.ToString() ?? Tokens.ByRef;

                if (param.IsObject || arg.GetDescendent<VBAParser.CtLExprContext>() != null || requiresByVal)
                {
                    modifier = Tokens.ByVal;
                }

                if (param.IsParamArray)
                {
                    modifier = Tokens.ParamArray;
                    return (string.Empty, arg.GetText().Replace($"{param.AsTypeName}", string.Empty).Trim(), param.AsTypeName);
                }
                return (modifier, param.IdentifierName, param.AsTypeName);
            }
        }
    }
}
