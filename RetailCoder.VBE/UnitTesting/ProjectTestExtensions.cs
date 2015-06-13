using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;

using Rubberduck.Parsing;
using Rubberduck.Parsing.Reflection;
using Rubberduck.Reflection;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.UnitTesting
{
    [ComVisible(false)]
    public static class ProjectTestExtensions
    {
        /// <summary>
        /// Runs all methods with specified attribute.
        /// </summary>
        /// <typeparam name="TAttribute"></typeparam>
        /// <param name="component"></param>
        /// <remarks>
        /// Order of execution cannot be garanteed.
        /// </remarks>
        public static void RunMethodsWithAttribute<TAttribute>(this VBComponent component)
            where TAttribute : MemberAttributeBase, new()
        {
            var hostApp = component.VBE.HostApplication();
            var methods = component.GetMembers(vbext_ProcKind.vbext_pk_Proc)
                                   .Where(member => member.HasAttribute<TAttribute>());
            foreach (var method in methods)
            {
                hostApp.Run(method.QualifiedMemberName);
            }
        }

        public static IEnumerable<TestMethod> TestMethods(this VBProject project)
        {
            var hostApp = project.VBE.HostApplication();

            var result = project.VBComponents
                          .Cast<VBComponent>()
                          .Where(component => component.CodeModule.HasAttribute<TestModuleAttribute>())
                          .Select(component => new { Component = component, Members = component.GetMembers().Where(IsTestMethod)})
                          .SelectMany(component => component.Members.Select(method => 
                              new TestMethod(method.QualifiedMemberName, hostApp)));

            return result;
        }

        public static IEnumerable<TestMethod> TestMethods(this VBComponent component)
        {
            var hostApp = component.VBE.HostApplication();

            if (component.Type == vbext_ComponentType.vbext_ct_StdModule 
                && component.CodeModule.HasAttribute<TestModuleAttribute>())
            {
                return component.GetMembers()
                                .Where(IsTestMethod)
                                .Select(member => new TestMethod(member.QualifiedMemberName, hostApp));
            }

            return new List<TestMethod>();
        }

        private static readonly string[] ReservedTestAttributeNames =
        {
            "ModuleInitialize",
            "TestInitialize", 
            "TestCleanup",
            "ModuleCleanup"
        };

        private static bool IsTestMethod(Member member)
        {
            var isIgnoredMethod = member.HasAttribute<TestInitializeAttribute>()
                               || member.HasAttribute<TestCleanupAttribute>()
                               || member.HasAttribute<ModuleInitializeAttribute>()
                               || member.HasAttribute<ModuleCleanupAttribute>()
                               || ReservedTestAttributeNames.Any(attribute => 
                                   member.QualifiedMemberName.MemberName.StartsWith(attribute));

            var result = !isIgnoredMethod &&
                (member.QualifiedMemberName.MemberName.StartsWith("Test") || member.HasAttribute<TestMethodAttribute>())
                 && member.Signature.Contains(member.QualifiedMemberName.MemberName + "()")
                 && member.MemberType == MemberType.Sub
                 && member.MemberVisibility == MemberVisibility.Public;

            return result;
        }
    }
}
