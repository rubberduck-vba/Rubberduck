using Antlr4.Runtime;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Parsing.Grammar
{
    // borrowed from https://stackoverflow.com/a/27217835
    internal class RecoveryStrategy : DefaultErrorStrategy
    {
        public override void Recover(Parser recognizer, RecognitionException e)
        {
            // This should should move the current position to the next 'END' token
            base.Recover(recognizer, e);

            ITokenStream tokenStream = (ITokenStream)recognizer.InputStream;

            // Verify we are where we expect to be
            if (tokenStream.La(1) == VBAParser.END)
            {
                // Get the next possible tokens
                IntervalSet intervalSet = GetErrorRecoverySet(recognizer);

                // Move to the next token
                tokenStream.Consume();

                // Move to the next possible token
                // If the errant element is the last in the set, this will move to the 'END' token in 'END MODULE'.
                // If there are subsequent elements in the set, this will move to the 'BEGIN' token in 'BEGIN module_element'.
                ConsumeUntil(recognizer, intervalSet);
            }
        }
    }
}
