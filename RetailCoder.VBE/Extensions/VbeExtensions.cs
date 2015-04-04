using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Nodes;
using Rubberduck.VBEHost;

namespace Rubberduck.Extensions
{
    public static class VbeExtensions
    {
        public static IEnumerable<CodeModule> FindCodeModules(this VBE vbe, QualifiedModuleName qualifiedName)
        {
            return FindCodeModules(vbe, qualifiedName.ProjectName, qualifiedName.ModuleName);
        }

        /// <summary>
        /// Finds all code modules that match the specified project and component names.
        /// </summary>
        /// <param name="vbe"></param>
        /// <param name="projectName"></param>
        /// <param name="componentName"></param>
        /// <returns></returns>
        public static IEnumerable<CodeModule> FindCodeModules(this VBE vbe, string projectName, string componentName)
        {
            var matches = 
                vbe.VBProjects.Cast<VBProject>()
                              .Where(project => project.Protection != vbext_ProjectProtection.vbext_pp_locked && project.Name == projectName)
                              .SelectMany(project => project.VBComponents.Cast<VBComponent>()
                                                                         .Where(component => component.Name == componentName))
                              .Select(component => component.CodeModule);
            return matches;
        }

        public static CodeModuleSelection FindInstruction(this VBE vbe, CommentNode comment)
        {
            var modules = FindCodeModules(vbe, comment.QualifiedSelection.QualifiedName);
            foreach (var module in modules)
            {
                var selection = comment.QualifiedSelection.Selection;

                if (module.Lines[selection.StartLine, selection.LineCount]
                    .Replace(" _\n", " ").Contains(comment.Comment))
                {
                    return new CodeModuleSelection(module, selection);
                }
            }

            return null;
        }

        public static void SetSelection(this VBE vbe, QualifiedSelection selection)
        {
            //not a very robust method. Breaks if there are multiple projects with the same name.
            var project = vbe.VBProjects.Cast<VBProject>()
                            .FirstOrDefault(p => p.Protection != vbext_ProjectProtection.vbext_pp_locked && p.Name == selection.QualifiedName.ProjectName);

            VBComponent component = null;
            if (project != null)
            {
                component = project.VBComponents.Cast<VBComponent>()
                                .FirstOrDefault(c => c.Name == selection.QualifiedName.ModuleName);
            }

            if (component == null)
            {
                return;
            }

            component.CodeModule.CodePane.SetSelection(selection.Selection);
        }

        [Obsolete]
        public static CodeModuleSelection FindInstruction(this VBE vbe, QualifiedModuleName qualifiedModuleName, ParserRuleContext context)
        {
            var projectName = qualifiedModuleName.ProjectName;
            var componentName = qualifiedModuleName.ModuleName;

            var modules = FindCodeModules(vbe, projectName, componentName).ToList();
            foreach (var module in modules)
            {
                Selection selection;
                var text = " ";
                if (context == null)
                {
                    selection = Selection.Home;
                }
                else
                {
                    selection = context.GetSelection();
                    text = context.GetText();
                }

                if (module.Lines[selection.StartLine, selection.LineCount]
                    .Replace(" _\n", " ").Contains(text))
                {
                    return new CodeModuleSelection(module, selection);
                }
            }

            return new CodeModuleSelection(modules.First(), Selection.Home);
        }

        public static CodeModuleSelection FindInstruction(this VBE vbe, QualifiedModuleName qualifiedModuleName, Selection selection)
        {
            var projectName = qualifiedModuleName.ProjectName;
            var componentName = qualifiedModuleName.ModuleName;

            var modules = FindCodeModules(vbe, projectName, componentName).ToList();

            return new CodeModuleSelection(modules.First(), selection);
        }

        /// <summary> Returns the type of Office Application that is hosting the VBE. </summary>
        /// <returns> Returns null if Unit Testing does not support Host Application.</returns>
        public static IHostApplication HostApplication(this VBE vbe)
        {
            foreach (Reference reference in vbe.ActiveVBProject.References)
            {
                if (reference.BuiltIn && reference.Name != "VBA")
                {
                    if (reference.Name == "Excel") return new ExcelApp();
                    if (reference.Name == "Access") return new AccessApp();
                    if (reference.Name == "Word") return new WordApp();
                    if (reference.Name == "PowerPoint") return new PowerPointApp();
                    if (reference.Name == "Outlook") return new OutlookApp();
                    if (reference.Name == "Publisher") return new PublisherApp();
                }
            }

            return null;
        }
    }
}