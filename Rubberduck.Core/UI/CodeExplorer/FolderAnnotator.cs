using System;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Annotations;

namespace Rubberduck.UI.CodeExplorer
{
    public static class FolderAnnotator
    {
        public static string AddOrUpdateFolderName((IParseTree parseTree, TokenStreamRewriter rewriter) startRule, string updatedFolderName)
        {
            if (HasModuleDeclarations(startRule.parseTree, out var moduleDeclarations))
            {
                if (HasFolderAnnotation(startRule.parseTree, out var folderAnnotation))
                {
                    var oldFolder = folderAnnotation.GetChild<VBAParser.AnnotationArgListContext>()
                        .GetChild<VBAParser.AnnotationArgContext>();
                    startRule.rewriter.Replace(oldFolder.SourceInterval.a, oldFolder.SourceInterval.b, updatedFolderName);
                }
                else
                {
                    var index = HasOptionExplicit(startRule.parseTree, out var optionExplicitStmt)
                        ? optionExplicitStmt.SourceInterval.a
                        : moduleDeclarations.SourceInterval.a;
                    startRule.rewriter.InsertBefore(index, FolderAnnotationWithFolderName(updatedFolderName) + Environment.NewLine);
                }
            }
            else
            {
                var moduleAttributes = ((ParserRuleContext)startRule.parseTree).GetDescendents<VBAParser.ModuleAttributesContext>().First();
                var lastModuleAttribute = moduleAttributes.GetChild(moduleAttributes.ChildCount - 1).GetText();
                var isLastAttributeNewLine = lastModuleAttribute.Equals(Environment.NewLine);
                if (isLastAttributeNewLine)
                {
                    startRule.rewriter.InsertBefore(moduleAttributes.SourceInterval.b, Environment.NewLine
                        + FolderAnnotationWithFolderName(updatedFolderName) + Environment.NewLine);
                }
                else
                {
                    startRule.rewriter.InsertAfter(moduleAttributes.SourceInterval.b, Environment.NewLine
                        + FolderAnnotationWithFolderName(updatedFolderName) + Environment.NewLine + Environment.NewLine);
                }
            }

            return startRule.rewriter.GetText();
        }

        private static bool HasModuleDeclarations(IParseTree parseTree, out VBAParser.ModuleDeclarationsContext moduleDeclarations)
        {
            var startRuleContext = (ParserRuleContext)parseTree;
            var moduleDescendents = startRuleContext.GetDescendents<VBAParser.ModuleDeclarationsContext>();
            if (!moduleDescendents.ElementAt(0).GetText().Equals(string.Empty))
            {
                moduleDeclarations = moduleDescendents.ElementAt(0);
                return true;
            }

            moduleDeclarations = null;
            return false;
        }

        private static bool HasFolderAnnotation(IParseTree parseTree, out VBAParser.AnnotationContext folderAnnotation)
        {
            var startRuleContext = (ParserRuleContext)parseTree;
            var folderDescendents = startRuleContext.GetDescendents<VBAParser.AnnotationContext>()
                                        .Where(a => a.GetText().Contains(AnnotationType.Folder.ToString()));
            if (folderDescendents.Any())
            {
                folderAnnotation = folderDescendents.ElementAt(0);
                return true;
            }

            folderAnnotation = null;
            return false;
        }

        private static bool HasOptionExplicit(IParseTree parseTree, out VBAParser.OptionExplicitStmtContext optionExplicit)
        {
            var startRuleContext = (ParserRuleContext)parseTree;
            var optionExplicitDescendents = startRuleContext.GetDescendents<VBAParser.OptionExplicitStmtContext>();
            if (optionExplicitDescendents.Any())
            {
                optionExplicit = optionExplicitDescendents.ElementAt(0);
                return true;
            }

            optionExplicit = null;
            return false;
        }

        private static string FolderAnnotationWithFolderName(string folderName)
        {
            return $"'@{AnnotationType.Folder}({folderName})";
        }
    }
}
