using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Reflection = System.Reflection;
using System.Linq;

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

        // returns the raw conditional arguments string, e.g. "foo = 1 : bar = 2"
        public static string GetProjectConditionalCompilationArgsRaw(IVBE ide, string projectName)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).ConditionalCompilationArguments;
            }
        }

        // return the parsed conditional arguments string as a Dictionary<string, string>
        public static Dictionary<string, string> GetProjectConditionalCompilationArgs(IVBE ide, string projectName)
        {
            var args = GetProjectConditionalCompilationArgsRaw(ide, projectName);

            if (args.Length > 0)
            { 
                string[] argsArray = args.Split(new[] { ':' });
                return argsArray.Select(item => item.Split('=')).ToDictionary(s => s[0], s => s[1]);
            }
            else
            {
                return new Dictionary<string, string>();
            }
        }

        // sets the raw conditional arguments string, e.g. "foo = 1 : bar = 2"
        public static void SetProjectConditionalCompilationArgsRaw(IVBE ide, string projectName, string newConditionalArgs)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                typeLibs.FindTypeLib(projectName).ConditionalCompilationArguments = newConditionalArgs;
            }
        }

        // sets the conditional arguments string via a Dictionary<string, string>
        public static void SetProjectConditionalCompilationArgs(IVBE ide, string projectName, Dictionary<string, string> newConditionalArgs)
        {
            var rawArgsString = string.Join(" : ", newConditionalArgs.Select(x => x.Key + " = " + x.Value));
            SetProjectConditionalCompilationArgsRaw(ide, projectName, rawArgsString);
        }

        public static bool IsExcelWorkbook(IVBE ide, string projectName, string className)
        {
            return DoesClassImplementInterface(ide, projectName, className, "Excel", "_Workbook");
        }

        public static bool IsExcelWorksheet(IVBE ide, string projectName, string className)
        {
            return DoesClassImplementInterface(ide, projectName, className, "Excel", "_Worksheet");
        }

        public static bool DoesClassImplementInterface(IVBE ide, string projectName, string className, string typeLibName, string interfaceName)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(className).DoesImplement(typeLibName, interfaceName);
            }
        }

        public static bool DoesClassImplementInterface(IVBE ide, string projectName, string className, Guid interfaceIID)
        {
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                return typeLibs.FindTypeLib(projectName).FindTypeInfo(className).DoesImplement(interfaceIID);
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
            using (var typeLibs = new TypeLibsAccessor_VBE(ide))
            {
                var documenter = new TypeLibDocumenter();

                foreach (var typeLib in typeLibs)
                {
                    documenter.AddTypeLib(typeLib);
                }

                return documenter.ToString();
            }            
        }
    }
}
