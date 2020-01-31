using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMovedMemberNewContentProvider
    {
        bool UsesExistingClassVariable { set; get; }
        ClassVariableContentProvider ClassVariableContentProvider { set; get; }
        void AddDeclarationBlock(string declarationBlock);
        void AddCodeBlock(string codeBlock);
        List<string> Declarations { get; }
        List<string> CodeBlocks { get; }
        string DestinationClassVariableDeclaration { get; }
        string DestinationClassInstantiationBlock { get; }
        string ClassInstantiationFragment { get; }
        string ClassInstantiationSubName { get; }
        string PreCodeSectionContent { get; }
        bool HasNewContent { get; }
        string AsCodeSection { get; }
        string AsNewModuleContent { get; }
        int CountOfProcLines { get; }
    }

    public class MoveMemberNewContent : IMovedMemberNewContentProvider
    {
        private ClassVariableContentProvider? _classVariableContentProvider;
        public ClassVariableContentProvider ClassVariableContentProvider
        {
            set
            {
                _classVariableContentProvider = value;
            }
            get
            {
                return _classVariableContentProvider.Value;
            }
        }

        private readonly MoveEndpoints _moveType;
        private IEnumerable<Declaration> SourceModuleElements { get; }

        public MoveMemberNewContent(MoveDefinition endpoints, IEnumerable<Declaration> sourceModueElements, ClassVariableContentProvider? classVariableContentProvider = null)
        {
            _moveType = endpoints.Endpoints;
            SourceModuleElements = sourceModueElements;
        }

        public bool UsesExistingClassVariable { set; get; } = false;

        public void AddDeclarationBlock(string declarationBlock)
        {
            if (declarationBlock.Length > 0)
            {
                Declarations.Add(declarationBlock);
            }
        }

        public void AddCodeBlock(string codeBlock)
        {
            if (codeBlock.Length > 0)
            {
                CodeBlocks.Add(codeBlock);
            }
        }

        private bool DestinationIsClassModule
            => _moveType == MoveEndpoints.ClassToClass || _moveType == MoveEndpoints.StdToClass || _moveType == MoveEndpoints.FormToClass;

        private bool DestinationIsStdModule => !DestinationIsClassModule;

        private bool SourceIsClassModule => _moveType == MoveEndpoints.ClassToClass || _moveType == MoveEndpoints.ClassToStd;

        private bool SourceIsStdModule => _moveType == MoveEndpoints.StdToClass || _moveType == MoveEndpoints.StdToStd;

        public List<string> Declarations { get; } = new List<string>();

        public List<string> CodeBlocks { get; } = new List<string>();

        public string DestinationClassVariableDeclaration
        {
            get
            {
                if (UsesExistingClassVariable)
                {
                    return string.Empty;
                }

                if (DestinationIsClassModule && _classVariableContentProvider.HasValue)
                {
                    return _classVariableContentProvider.Value.ClassVariableDeclaration;
                }
                return string.Empty;
            }
        }

        public string DestinationClassInstantiationBlock
        {
            get
            {
                if (DestinationIsStdModule || !_classVariableContentProvider.HasValue)
                {
                    return string.Empty;
                }

                if (SourceIsClassModule)
                {
                    var classInitialize = SourceModuleElements.FirstOrDefault(el => el.IdentifierName.Equals("Class_Initialize"));

                    return classInitialize is null ?
                        _classVariableContentProvider.Value.ClassModuleClassInitializeProcedure
                        : string.Empty;
                }
                return SourceModuleElements.FirstOrDefault(el => el.IdentifierName.StartsWith(_classVariableContentProvider.Value.StdModuleClassVariableInstantiationSubName)) is null
                ? _classVariableContentProvider.Value.StdModuleClassVariableInstantiationProcedure
                    : string.Empty;
            }
        }

        public string ClassInstantiationFragment
            => _classVariableContentProvider.HasValue ?
                _classVariableContentProvider.Value.ClassInstantiationFragment
                : string.Empty;

        public string ClassInstantiationSubName
            => _classVariableContentProvider.HasValue ?
                _classVariableContentProvider.Value.StdModuleClassVariableInstantiationSubName
                : string.Empty;

        public string PreCodeSectionContent
        {
            get
            {
                var preCodeSectionContent = new List<string>(Declarations);
                if (DestinationClassVariableDeclaration.Length > 0)
                {
                    preCodeSectionContent.Add(DestinationClassVariableDeclaration);
                }
                if (DestinationClassInstantiationBlock.Length > 0)
                {
                    preCodeSectionContent.Add(DestinationClassInstantiationBlock);
                }
                if (preCodeSectionContent.Any())
                {
                    var preCodeSection = string.Join(Environment.NewLine, preCodeSectionContent);
                    return $"{Environment.NewLine}{preCodeSection}{Environment.NewLine}{Environment.NewLine}";
                }
                return string.Empty;
            }
        }

        public bool HasNewContent
            => Declarations.Any() || CodeBlocks.Any() || DestinationClassVariableDeclaration.Length > 0 || DestinationClassInstantiationBlock.Length > 0;

        public string AsCodeSection => CodeBlocks.Any() ? string.Join($"{Environment.NewLine}{Environment.NewLine}", CodeBlocks) : string.Empty;

        public string AsNewModuleContent
            => HasNewContent ? $"{PreCodeSectionContent}{Environment.NewLine}{Environment.NewLine}{AsCodeSection}" : string.Empty;

        public int CountOfProcLines => AsCodeSection.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Count();
    }
}
