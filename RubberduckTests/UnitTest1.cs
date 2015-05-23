using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using LibGit2Sharp;

namespace RubberduckTests
{
    [TestClass]
    public class UnitTest1
    {
        readonly string _repoPath =
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SourceControlTest");

        [TestMethod]
        public void TestMethod1()
        {
            
            using (var repo = new Repository(_repoPath))
            {
                //var filter = new CommitFilter()
                //{
                //    Since = repo.Branches["master"],
                //    Until = repo.Branches["origin/master"]
                //};

                //var commits = repo.Commits.QueryBy(filter).ToList();

                //var local = repo.Branches["master"].Commits;
                var currentBranch = repo.Branches["master"];
                var local = currentBranch.Commits;

                //var remote = repo.Branches["origin/master"].Commits;
                ICommitLog remote;
                List<Commit> localUnsynced;
                List<Commit> remoteUnsycned;

                if (currentBranch.TrackedBranch == null)
                {
                    localUnsynced = local.ToList();
                    remoteUnsycned = new List<Commit>();
                }
                else
                {
                    remote = currentBranch.TrackedBranch.Commits;

                    localUnsynced = local.Where(c => !remote.Contains(c)).ToList();
                    remoteUnsycned = remote.Where(c => !local.Contains(c)).ToList();
                }

               

            }

        }
    }
}