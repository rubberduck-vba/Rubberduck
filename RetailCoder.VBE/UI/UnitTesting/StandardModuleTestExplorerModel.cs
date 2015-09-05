using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Reflection;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.UnitTesting
{
    /// <summary>
    /// A TestExplorer model that discovers unit tests in standard modules (.bas) marked with a '@TestModule marker.
    /// </summary>
    public class StandardModuleTestExplorerModel : TestExplorerModelBase
    {
        private readonly VBE _vbe;

        public StandardModuleTestExplorerModel(VBE vbe)
        {
            _vbe = vbe;
        }

        public override void Refresh()
        {
            var tests = _vbe.VBProjects.Cast<VBProject>()
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                .SelectMany(project => project.VBComponents.Cast<VBComponent>())
                .Where(component => component.CodeModule.HasAttribute<TestModuleAttribute>())
                .Select(component => new { Component = component, Members = component.GetMembers(vbext_ProcKind.vbext_pk_Proc).Where(IsTestMethod) })
                .SelectMany(component => component.Members.Select(method =>
                    new TestMethod(method.QualifiedMemberName, _vbe)));

            Tests.Clear();
            foreach (var test in tests)
            {                
                Tests.Add(test);
            }

            OnPropertyChanged("Tests");
        }
    }
}