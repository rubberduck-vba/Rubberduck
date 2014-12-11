using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Extensions;

using Microsoft.Vbe.Interop;
using Rubberduck.Reflection;

namespace Rubberduck.UnitTesting
{
    internal static class ProjectTestExtensions
    {
        public static IEnumerable<VBComponent> TestModules(this VBProject project)
        {
            return project.VBComponents
                          .Cast<VBComponent>()
                          .Where(component => component.CodeModule.HasAttribute<TestModuleAttribute>());
        }

        public static IEnumerable<TestMethod> TestMethods(this VBProject project)
        {
            IHostApplication hostApp = project.VBE.HostApplication();

            return project.VBComponents
                          .Cast<VBComponent>()
                          .Where(component => component.Type == vbext_ComponentType.vbext_ct_StdModule && component.CodeModule.HasAttribute<TestModuleAttribute>())
                          .Select(component => new { Component = component, Members = component.GetMembers().Where(member => IsTestMethod(member))})
                          .SelectMany(component => component.Members.Select(method => new TestMethod(project.Name, component.Component.Name, method.Name, hostApp)));
        }

        public static IEnumerable<TestMethod> TestMethods(this VBComponent component)
        {
            IHostApplication hostApp = component.VBE.HostApplication();

            if (component.Type == vbext_ComponentType.vbext_ct_StdModule && component.CodeModule.HasAttribute<TestModuleAttribute>())
            {
                return component.GetMembers().Where(member => IsTestMethod(member))
                                .Select(member => new TestMethod(component.Collection.Parent.Name, component.Name, member.Name, hostApp));
            }

            return new List<TestMethod>();
        }

        private static bool IsTestMethod(Member member)
        {
            return (member.Name.StartsWith("Test") || member.HasAttribute<TestMethodAttribute>())
                 && member.Signature.Contains(member.Name + "()")
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
