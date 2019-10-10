using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

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
            return ArrangeIndentCurrentModuleCommand(vbe, state, indenter, MockVbeEvents.CreateMockVbeEvents(vbe));
        }

        internal static IndentCurrentModuleCommand ArrangeIndentCurrentModuleCommand(Mock<IVBE> vbe, RubberduckParserState state, IIndenter indenter, Mock<IVbeEvents> vbeEvents)
        {
            return new IndentCurrentModuleCommand(vbe.Object, indenter, state, vbeEvents.Object);
        }

        internal static NoIndentAnnotationCommand ArrangeNoIndentAnnotationCommand(Mock<IVBE> vbe,
            RubberduckParserState state)
        {
            return ArrangeNoIndentAnnotationCommand(vbe, state, MockVbeEvents.CreateMockVbeEvents(vbe));
        }

        internal static NoIndentAnnotationCommand ArrangeNoIndentAnnotationCommand(Mock<IVBE> vbe,
            RubberduckParserState state, Mock<IVbeEvents> vbeEvents)
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
            return ArrangeIndentCurrentProcedureCommand(vbe, indenter, state, MockVbeEvents.CreateMockVbeEvents(vbe));
        }

        internal static IndentCurrentProcedureCommand ArrangeIndentCurrentProcedureCommand(Mock<IVBE> vbe,
            IIndenter indenter, RubberduckParserState state, Mock<IVbeEvents> vbeEvents)
        {
            return new IndentCurrentProcedureCommand(vbe.Object, indenter, state, vbeEvents.Object);
        }
    }
}