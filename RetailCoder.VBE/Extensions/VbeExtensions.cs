using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Extensions
{
    [ComVisible(false)]
    public static class VbeExtensions
    {
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
                              .Where(project => project.Name == projectName)
                              .SelectMany(project => project.VBComponents.Cast<VBComponent>()
                                                                         .Where(component => component.Name == componentName))
                              .Select(component => component.CodeModule);
            return matches;
        }

        public static CodeModuleSelection FindInstruction(this VBE vbe, Instruction instruction)
        {
            var projectName = instruction.Line.ProjectName;
            var componentName = instruction.Line.ComponentName;

            var modules = FindCodeModules(vbe, projectName, componentName);
            foreach (var module in modules)
            {
                var startLine = instruction.Selection.StartLine == 0 ? 1 : instruction.Selection.StartLine;

                if (module.Lines[startLine, instruction.Selection.LineCount]
                         .Replace("_", string.Empty)
                         .Replace("\n\r", string.Empty).Contains(instruction.Content))
                {
                    return new CodeModuleSelection(module, instruction.Selection);
                }
            }

            return null;
        }

        /// <summary> Returns the type of Office Application that is hosting the VBE. </summary>
        public static HostApplicationType HostApplication(this VBE vbe)
        {
            foreach (Reference reference in vbe.ActiveVBProject.References)
            {
                if (reference.BuiltIn && reference.Name != "VBA")
                {
                    if (reference.Name == "Excel") return HostApplicationType.Excel;
                    if (reference.Name == "Access") return HostApplicationType.Access;
                    if (reference.Name == "Word") return HostApplicationType.Word;
                }
            }

            return HostApplicationType.Unknown;
        }
    }
}