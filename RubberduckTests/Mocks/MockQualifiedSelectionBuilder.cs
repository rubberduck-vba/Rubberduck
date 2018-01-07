using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Mocks
{
    public class MockQualifiedSelectionBuilder
    {
        private readonly IVBComponent _component;
        public MockQualifiedSelectionBuilder(IVBComponent component)
        {
            _component = component;
        }

        public QualifiedSelection CreateQualifiedSelection(Selection selection)
        {
            return new QualifiedSelection(new QualifiedModuleName(_component), selection);
        }

        public static QualifiedSelection CreateQualifiedSelection(IVBComponent component, Selection selection)
        {
            var builder = new MockQualifiedSelectionBuilder(component);
            return builder.CreateQualifiedSelection(selection);
        }
    }
}
