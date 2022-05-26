using NUnit.Framework;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.DeleteDeclarations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.DeleteDeclarations
{
    [TestFixture]
    public class NonDeleteIndicePairGeneratorTests
    {
        private  readonly DeleteDeclarationsTestSupport _support = new DeleteDeclarationsTestSupport();

        [TestCase("0", "0:-1")]
        [TestCase("1", "-1:1,1:-1")]
        [TestCase("1,5", "-1:1,1:5,5:-1")]
        [TestCase("1,5,6,7", "-1:1,1:5,7:-1")]
        [TestCase("0,3,5", "0:3,3:5,5:-1")]
        [TestCase("1,5,6,7,10,11", "-1:1,1:5,7:10,11:-1")]
        [TestCase("1,2,4,5,6,9", "-1:1,2:4,6:9,9:-1")]
        [Category("Refactorings")]
        [Category(nameof(DeleteDeclarationsRefactoringAction))]
        public void NonDeleteIndicePairGenerationTests(string inputIndices, string expectedPairs)
        {
            var nonDeleteIndices = new List<int>();
            DelimitedTokensToList(inputIndices, ",").ForEach(t => nonDeleteIndices.Add(int.Parse(t)));

            var results = NonDeleteIndicePairGenerator.Generate(nonDeleteIndices);

            var expected = new List<(int?, int?)>();
            var pairs = DelimitedTokensToList(expectedPairs, ",");
            foreach (var pair in pairs)
            {
                var vals = DelimitedTokensToList(pair, ":");

                expected.Add((ParseToNullable(vals[0]), ParseToNullable(vals[1])));
            }

            foreach (var exp in expected)
            {
                Assert.Contains(exp, results);
            }
        }

        private int? ParseToNullable(string token)
        {
            var val = int.Parse(token);
            return val < 0 ? null : val as int?;
        }

        private List<string> DelimitedTokensToList(string input, string delimiter)
        {
            return input.Split(new string[] { delimiter }, StringSplitOptions.None).ToList();
        }
    }
}
