using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Reflection = System.Reflection;

namespace Rubberduck.VBEditor.ComManagement.TypeLibsAPI
{
    public class VBETypeLibsAPI
    {
        public static void ExecuteCode(IVBE ide, string projectName, string standardModuleName, string procName, object[] args = null)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                typeLibs.FindTypeLib(projectName).FindTypeInfo(standardModuleName)
                    .StdModExecute(procName, Reflection.BindingFlags.InvokeMethod, args);
            }
        }

        public static string GetProjectConditionalCompilationArgs(IVBE ide, string projectName)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).ConditionalCompilationArguments;
            }
        }

        public static void SetProjectConditionalCompilationArgs(IVBE ide, string projectName, string newConditionalArgs)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                typeLibs.FindTypeLib(projectName).ConditionalCompilationArguments = newConditionalArgs;
            }
        }

        public static bool IsAWorkbook(IVBE ide, string projectName, string className)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(className).DoesImplement("_Workbook");
            }
        }

        public static bool IsAWorksheet(IVBE ide, string projectName, string className)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(className).DoesImplement("_Worksheet");
            }
        }

        public static string GetUserFormControlType(IVBE ide, string projectName, string userFormName, string controlName)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(userFormName)
                        .GetImplementedTypeInfo("FormItf").GetControlType(controlName).Name;
            }
        }

        public static string DocumentAll(IVBE ide)
        {
            var documenter = new TypeLibDocumenter();

            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                foreach (var typeLib in typeLibs)
                {
                    documenter.AddTypeLib(typeLib);
                }
            }

            return documenter.ToString();
        }
    }
}
