using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Commands
{
    public class MockIndenter
    {
        internal static IIndenter CreateIndenter(IVBE vbe)
        {
            return new Indenter(vbe, () => Settings.IndenterSettingsTests.GetMockIndenterSettings());
        }

        internal static IndentCurrentModuleCommand ArrangeIndentCurrentModuleCommand(Mock<IVBE> vbe,
            RubberduckParserState state)
        {
            return ArrangeIndentCurrentModuleCommand(vbe, state, CreateIndenter(vbe.Object));
        }

        internal static IndentCurrentModuleCommand ArrangeIndentCurrentModuleCommand(Mock<IVBE> vbe,
            RubberduckParserState state, IIndenter indenter)
        {
            return ArrangeIndentCurrentModuleCommand(vbe, state, indenter, new Mock<IVBEEvents>());
        }

        internal static IndentCurrentModuleCommand ArrangeIndentCurrentModuleCommand(Mock<IVBE> vbe, RubberduckParserState state, IIndenter indenter, Mock<IVBEEvents> vbeEvents)
        {
            return new IndentCurrentModuleCommand(vbe.Object, indenter, state, vbeEvents.Object);
        }

        internal static NoIndentAnnotationCommand ArrangeNoIndentAnnotationCommand(Mock<IVBE> vbe,
            RubberduckParserState state)
        {
            return ArrangeNoIndentAnnotationCommand(vbe, state, new Mock<IVBEEvents>());
        }

        internal static NoIndentAnnotationCommand ArrangeNoIndentAnnotationCommand(Mock<IVBE> vbe,
            RubberduckParserState state, Mock<IVBEEvents> vbeEvents)
        {
            return new NoIndentAnnotationCommand(vbe.Object, state, vbeEvents.Object);
        }

        internal static IndentCurrentProcedureCommand ArrangeIndentCurrentProcedureCommand(Mock<IVBE> vbe,
            RubberduckParserState state)
        {
            return ArrangeIndentCurrentProcedureCommand(vbe, CreateIndenter(vbe.Object), state);
        }

        internal static IndentCurrentProcedureCommand ArrangeIndentCurrentProcedureCommand(Mock<IVBE> vbe,
            IIndenter indenter, RubberduckParserState state)
        {
            return ArrangeIndentCurrentProcedureCommand(vbe, indenter, state, new Mock<IVBEEvents>());
        }

        internal static IndentCurrentProcedureCommand ArrangeIndentCurrentProcedureCommand(Mock<IVBE> vbe,
            IIndenter indenter, RubberduckParserState state, Mock<IVBEEvents> vbeEvents)
        {
            return new IndentCurrentProcedureCommand(vbe.Object, indenter, state, vbeEvents.Object);
        }
    }
}