using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Reflection;
using Rubberduck.VBA;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.UnitTesting
{
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

            return project.VBComponents
                          .Cast<VBComponent>()
                          .Where(component => component.CodeModule.HasAttribute<TestModuleAttribute>())
                          .Select(component => new { Component = component, Members = component.GetMembers().Where(IsTestMethod)})
                          .SelectMany(component => component.Members.Select(method => 
                              new TestMethod(method.QualifiedMemberName, hostApp)));
        }

        public static IEnumerable<TestMethod> TestMethods(this VBComponent component)
        {
            IHostApplication hostApp = component.VBE.HostApplication();

            if (component.Type == vbext_ComponentType.vbext_ct_StdModule 
                && component.CodeModule.HasAttribute<TestModuleAttribute>())
            {
                return component.GetMembers()
                                .Where(IsTestMethod)
                                .Select(member => new TestMethod(member.QualifiedMemberName, hostApp));
            }

            return new List<TestMethod>();
        }

        private static readonly string[] ReservedTestAttributeNames = {"TestInitialize", "TestCleanup"};

        private static bool IsTestMethod(Member member)
        {
            return (member.QualifiedMemberName.MemberName.StartsWith("Test") || member.HasAttribute<TestMethodAttribute>())
                 && member.Signature.Contains(member.QualifiedMemberName.MemberName + "()")
                 && !ReservedTestAttributeNames.Contains(member.QualifiedMemberName.MemberName)
                 && member.MemberType == MemberType.Sub
                 && member.MemberVisibility == MemberVisibility.Public;
        }

        private static bool IsTestModule(CodeModule module)
        {
            return (module.Parent.Type == vbext_ComponentType.vbext_ct_StdModule
                && module.Name.StartsWith("Test") || module.HasAttribute<TestModuleAttribute>());
        }
    }
}
