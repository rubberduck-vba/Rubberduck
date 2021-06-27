using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    internal class NonDeleteIndicePairGenerator
    {
        /// <summary>
        /// Given a list of non-negative integers, generates a list of Nullable&lt;int&gt; pairs 
        /// that are the bounding values of missing integer values.  
        /// e.g., 1,2,4,5,6,9 generates (null, 1), (2, 4), (6, 9), (9, null)
        /// </summary>
        /// <remarks>
        /// A null first value implies all indices up to the second value of the pair.
        /// A null second value implies all indices following the first value of the pair.
        /// </remarks>
        public static List<(int?, int?)> Generate(List<int> nonDeleteIndices)
        {
            var results = new List<(int?, int?)>();

            if (!nonDeleteIndices.Any())
            {
                return results;
            }

            var startingNonDeleteContextIndex = nonDeleteIndices.ElementAt(0);
            if (startingNonDeleteContextIndex != 0)
            {
                results.Add((null, startingNonDeleteContextIndex));
            }

            for (var nonDeleteIndex = 0; nonDeleteIndex + 1 < nonDeleteIndices.Count; nonDeleteIndex++)
            {
                var currentIndice = nonDeleteIndices.ElementAt(nonDeleteIndex);
                var nextIndice = nonDeleteIndices.ElementAt(nonDeleteIndex + 1);
                if (nonDeleteIndices.ElementAt(nonDeleteIndex + 1) - nonDeleteIndices.ElementAt(nonDeleteIndex) == 1)
                {
                    continue;
                }
                results.Add((currentIndice, nextIndice));
            }

            results.Add((nonDeleteIndices.Last(), null));
            return results;
        }
    }
}
