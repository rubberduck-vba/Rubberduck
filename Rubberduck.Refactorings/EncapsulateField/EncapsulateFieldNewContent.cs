using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldNewContentProvider
    {
        void AddDeclarationBlock(string declarationBlock);
        void AddCodeBlock(string codeBlock);
        List<string> Declarations { get; }
        List<string> CodeBlocks { get; }
        string PreCodeSectionContent { get; }
        bool HasNewContent { get; }
        string CodeSectionContent { get; }
        string AsSingleTextBlock { get; }
        int CountOfProcLines { get; }
    }

    public class EncapsulateFieldNewContent : IEncapsulateFieldNewContentProvider
    {
        protected static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";
        protected static string SingleSpace => $"{Environment.NewLine}";

        private IEnumerable<Declaration> SourceModuleElements { get; }

        public void AddTypeDeclarationBlock(string declarationBlock)
        {
            if (declarationBlock.Length > 0)
            {
                TypeDeclarations.Add(declarationBlock);
            }
        }

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

        private string _postPendComment;
        public string PostPendComment
        {
            get => _postPendComment;
            set
            {
                if (value.Length > 0)
                {
                    _postPendComment = value.StartsWith("'") ? value : $"'{value}";
                }
            }
        }

        public List<string> TypeDeclarations { get; } = new List<string>();

        public List<string> Declarations { get; } = new List<string>();

        public List<string> CodeBlocks { get; } = new List<string>();

        public string PreCodeSectionContent
        {
            get
            {
                var allDeclarations = Enumerable.Empty<string>().Concat(TypeDeclarations).Concat(Declarations);
                var preCodeSectionContent = new List<string>(allDeclarations);
                if (preCodeSectionContent.Any())
                {
                    var preCodeSection = string.Join(SingleSpace, preCodeSectionContent);
                    return preCodeSection;
                }
                return string.Empty;
            }
        }

        public bool HasNewContent => Declarations.Any() || CodeBlocks.Any();

        public string CodeSectionContent => CodeBlocks.Any() ? string.Join($"{DoubleSpace}", CodeBlocks) : string.Empty;

        public string AsSingleTextBlock
        {
            get
            {
                if (!HasNewContent) { return string.Empty; }

                var content = string.Empty;
                if (PreCodeSectionContent.Length > 0)
                {
                    content = CodeSectionContent.Length > 0 
                        ? $"{PreCodeSectionContent}{DoubleSpace}{CodeSectionContent}"
                        : $"{PreCodeSectionContent}";
                }
                else
                {
                    content = CodeSectionContent.Length > 0
                        ? $"{SingleSpace}{CodeSectionContent}"
                        : string.Empty;
                }

                if (PostPendComment != null && PostPendComment.Length > 0)
                {
                    content = $"{content}{SingleSpace}{PostPendComment}{SingleSpace}";
                }
                return content;
            }
        }

        public int CountOfProcLines => CodeSectionContent.Split(new string[] {SingleSpace}, StringSplitOptions.None).Count();
    }
}
