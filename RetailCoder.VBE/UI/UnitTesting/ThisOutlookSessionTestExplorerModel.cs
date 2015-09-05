using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.UnitTesting
{
    /// <summary>
    /// A TestExplorer model that discovers unit tests in a 'ThisOutlookSession' document/class module.
    /// </summary>
    /// <remarks>
    /// We can *discover* unit test methods all we want... we can't run them.
    /// </remarks>
    public class ThisOutlookSessionTestExplorerModel : TestExplorerModelBase
    {
        private readonly VBE _vbe;

        public ThisOutlookSessionTestExplorerModel(VBE vbe)
        {
            _vbe = vbe;
        }

        public override void Refresh()
        {
            Tests.Clear();
            var tests = _vbe.ActiveVBProject.VBComponents.Cast<VBComponent>()
                .SingleOrDefault(component => component.Type == vbext_ComponentType.vbext_ct_Document)
                .GetMembers(vbext_ProcKind.vbext_pk_Proc).Where(IsTestMethod)
                .Select(method => new TestMethod(method.QualifiedMemberName, _vbe));

            foreach (var test in tests)
            {
                Tests.Add(test);
            }
        }
    }
}