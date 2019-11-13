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
        private IEnumerable<Declaration> SourceModuleElements { get; }

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

        public List<string> Declarations { get; } = new List<string>();

        public List<string> CodeBlocks { get; } = new List<string>();

        public string PreCodeSectionContent
        {
            get
            {
                var preCodeSectionContent = new List<string>(Declarations);
                if (preCodeSectionContent.Any())
                {
                    var preCodeSection = string.Join(Environment.NewLine, preCodeSectionContent);
                    return preCodeSection;
                }
                return string.Empty;
            }
        }

        public bool HasNewContent => Declarations.Any() || CodeBlocks.Any();

        public string CodeSectionContent => CodeBlocks.Any() ? string.Join($"{Environment.NewLine}{Environment.NewLine}", CodeBlocks) : string.Empty;

        public string AsSingleTextBlock
        {
            get
            {
                if (!HasNewContent) { return string.Empty; }

                if (PreCodeSectionContent.Length > 0)
                {
                    return CodeSectionContent.Length > 0 
                        ? $"{PreCodeSectionContent}{Environment.NewLine}{Environment.NewLine}{CodeSectionContent}"
                        : $"{PreCodeSectionContent}";
                }

                return $"{Environment.NewLine}{CodeSectionContent}";
            }
        }

        public int CountOfProcLines => CodeSectionContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Count();
    }
}
