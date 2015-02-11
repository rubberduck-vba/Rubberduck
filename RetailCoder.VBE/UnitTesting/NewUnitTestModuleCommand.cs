using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Reflection;
using Rubberduck.Extensions;

namespace Rubberduck.UnitTesting
{
    [ComVisible(false)]
    public static class NewUnitTestModuleCommand
    {
        private static readonly string TestModuleEmptyTemplate = string.Concat(
            "'@TestModule\n",
            "Private Assert As New Rubberduck.AssertClass\n\n"
            );

        private static readonly string TestModuleBaseName = "TestModule";

        public static void NewUnitTestModule(VBE vbe)
        {
            var project = vbe.ActiveVBProject;
            project.EnsureReferenceToAddInLibrary();

            var module = project.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            module.Name = GetNextTestModuleName(project);

            var hasOptionExplicit = false;
            if (module.CodeModule.CountOfLines > 0 && module.CodeModule.CountOfDeclarationLines > 0)
            {
                hasOptionExplicit = module.CodeModule.Lines[1, module.CodeModule.CountOfDeclarationLines].Contains("Option Explicit");
            }

            var options = string.Concat(hasOptionExplicit ? string.Empty : "Option Explicit\n", "Option Private Module\n\n");

            module.CodeModule.AddFromString(options + TestModuleEmptyTemplate);
            module.Activate();
        }

        private static string GetNextTestModuleName(VBProject project)
        {
            var names = project.ComponentNames();
            var index = names.Count(n => n.StartsWith(TestModuleBaseName)) + 1;

            return string.Concat(TestModuleBaseName, index);
        }
    }
}
