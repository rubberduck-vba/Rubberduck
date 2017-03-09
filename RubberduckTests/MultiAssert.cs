using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace RubberduckTests
{
    // borrowed from https://blog.elgaard.com/2011/02/06/multiple-asserts-in-a-single-unit-test-method/
    public static class MultiAssert
    {
        public static void Aggregate(params Action[] actions)
        {
            var exceptions = new List<AssertFailedException>();

            foreach (var action in actions)
            {
                try
                {
                    action();
                }
                catch (AssertFailedException ex)
                {
                    exceptions.Add(ex);
                }
            }

            var assertionTexts =
                exceptions.Select(assertFailedException => assertFailedException.Message).ToList();
            if (0 != assertionTexts.Count)
            {
                throw new
                    AssertFailedException(
                    assertionTexts.Aggregate(
                        (aggregatedMessage, next) => aggregatedMessage + Environment.NewLine + next));
            }
        }
    }
}
