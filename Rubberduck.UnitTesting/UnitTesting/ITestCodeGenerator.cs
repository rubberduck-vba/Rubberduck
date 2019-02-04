using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting.UnitTesting
{
    public interface ITestCodeGenerator
    {
        void AddTestModuleToProject(IVBProject project);
        void AddTestModuleToProject(IVBProject project, Declaration stubSource);
        string GetNewTestMethodCode(IVBComponent component);
        string GetNewTestMethodCodeErrorExpected(IVBComponent component);
    }
}
