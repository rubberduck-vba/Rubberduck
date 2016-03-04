using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class OptionCompareListener : VBABaseListener
    {
        public VBAOptionCompare OptionCompare { get; set; }

        public OptionCompareListener()
        {
            // Default is binary according to VBA specification.
            OptionCompare = VBAOptionCompare.Binary;
        }

        public override void ExitOptionCompareStmt([NotNull] VBAParser.OptionCompareStmtContext context)
        {
            if (context.TEXT() != null || context.DATABASE() != null)
            {
                OptionCompare = VBAOptionCompare.Text;
            }
        }
    }
}
