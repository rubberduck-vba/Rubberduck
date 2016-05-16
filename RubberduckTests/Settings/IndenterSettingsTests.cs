using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Settings;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class IndenterSettingsTests
    {
        private Configuration GetDefaultConfig()
        {
            var indenterSettings = new Rubberduck.Settings.IndenterSettings
            {
                IndentEntireProcedureBody = true,
                IndentFirstCommentBlock = true,
                IndentFirstDeclarationBlock = true,
                AlignCommentsWithCode = true,
                AlignContinuations = true,
                IgnoreOperatorsInContinuations = true,
                IndentCase = false,
                ForceDebugStatementsInColumn1 = false,
                ForceCompilerDirectivesInColumn1 = false,
                IndentCompilerDirectives = true,
                AlignDims = false,
                AlignDimColumn = 15,
                EnableUndo = true,
                EndOfLineCommentStyle = Rubberduck.SmartIndenter.EndOfLineCommentStyle.AlignInColumn,
                EndOfLineCommentColumnSpaceAlignment = 50,
                IndentSpaces = 4
            };

            var userSettings = new UserSettings(null, null, null, null, indenterSettings);
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            var indenterSettings = new Rubberduck.Settings.IndenterSettings
            {
                IndentEntireProcedureBody = false,
                IndentFirstCommentBlock = false,
                IndentFirstDeclarationBlock = false,
                AlignCommentsWithCode = false,
                AlignContinuations = false,
                IgnoreOperatorsInContinuations = false,
                IndentCase = true,
                ForceDebugStatementsInColumn1 = true,
                ForceCompilerDirectivesInColumn1 = true,
                IndentCompilerDirectives = false,
                AlignDims = true,
                AlignDimColumn = 16,
                EnableUndo = false,
                EndOfLineCommentStyle = Rubberduck.SmartIndenter.EndOfLineCommentStyle.Absolute,
                EndOfLineCommentColumnSpaceAlignment = 60,
                IndentSpaces = 2
            };

            var userSettings = new UserSettings(null, null, null, null, indenterSettings);
            return new Configuration(userSettings);
        }

        [TestMethod]
        public void SaveConfigWorks()
        {
            var customConfig = GetNondefaultConfig();
            var viewModel = new IndenterSettingsViewModel(customConfig);

            var config = GetDefaultConfig();
            viewModel.UpdateConfig(config);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.AlignCommentsWithCode),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.AlignContinuations, viewModel.AlignContinuations),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.AlignDimColumn, viewModel.AlignDimColumn),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.AlignDims, viewModel.AlignDims),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.EnableUndo, viewModel.EnableUndo),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.EndOfLineCommentColumnSpaceAlignment),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.EndOfLineCommentStyle),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.ForceCompilerDirectivesInColumn1),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.ForceDebugStatementsInColumn1),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.IgnoreOperatorsInContinuations),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentCase, viewModel.IndentCase),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.IndentCompilerDirectives),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.IndentEntireProcedureBody),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.IndentFirstCommentBlock),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.IndentFirstDeclarationBlock),
                () => Assert.AreEqual(config.UserSettings.IndenterSettings.IndentSpaces, viewModel.IndentSpaces));
        }

        [TestMethod]
        public void SetDefaultsWorks()
        {
            var viewModel = new IndenterSettingsViewModel(GetNondefaultConfig());

            var defaultConfig = GetDefaultConfig();
            viewModel.SetToDefaults(defaultConfig);

            MultiAssert.Aggregate(
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.AlignCommentsWithCode),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignContinuations, viewModel.AlignContinuations),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDimColumn, viewModel.AlignDimColumn),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDims, viewModel.AlignDims),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EnableUndo, viewModel.EnableUndo),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.EndOfLineCommentColumnSpaceAlignment),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.EndOfLineCommentStyle),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.ForceCompilerDirectivesInColumn1),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.ForceDebugStatementsInColumn1),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.IgnoreOperatorsInContinuations),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCase, viewModel.IndentCase),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.IndentCompilerDirectives),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.IndentEntireProcedureBody),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.IndentFirstCommentBlock),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.IndentFirstDeclarationBlock),
                () => Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentSpaces, viewModel.IndentSpaces));
        }

        [TestMethod]
        public void AlignCommentsWithCodeIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignCommentsWithCode, viewModel.AlignCommentsWithCode);
        }

        [TestMethod]
        public void AlignContinuationsIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignContinuations, viewModel.AlignContinuations);
        }

        [TestMethod]
        public void AlignDimColumnIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDimColumn, viewModel.AlignDimColumn);
        }

        [TestMethod]
        public void AlignDimsIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.AlignDims, viewModel.AlignDims);
        }

        [TestMethod]
        public void EnableUndoIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EnableUndo, viewModel.EnableUndo);
        }

        [TestMethod]
        public void EndOfLineCommentColumnSpaceAlignmentIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentColumnSpaceAlignment, viewModel.EndOfLineCommentColumnSpaceAlignment);
        }

        [TestMethod]
        public void EndOfLineCommentStyleIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.EndOfLineCommentStyle, viewModel.EndOfLineCommentStyle);
        }

        [TestMethod]
        public void ForceCompilerDirectivesInColumn1IsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceCompilerDirectivesInColumn1, viewModel.ForceCompilerDirectivesInColumn1);
        }

        [TestMethod]
        public void ForceDebugStatementsInColumn1IsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.ForceDebugStatementsInColumn1, viewModel.ForceDebugStatementsInColumn1);
        }

        [TestMethod]
        public void IgnoreOperatorsInContinuationsIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IgnoreOperatorsInContinuations, viewModel.IgnoreOperatorsInContinuations);
        }

        [TestMethod]
        public void IndentCaseIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCase, viewModel.IndentCase);
        }

        [TestMethod]
        public void IndentCompilerDirectivesIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentCompilerDirectives, viewModel.IndentCompilerDirectives);
        }

        [TestMethod]
        public void IndentEntireProcedureBodyIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentEntireProcedureBody, viewModel.IndentEntireProcedureBody);
        }

        [TestMethod]
        public void IndentFirstCommentBlockIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstCommentBlock, viewModel.IndentFirstCommentBlock);
        }

        [TestMethod]
        public void IndentFirstDeclarationBlockIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentFirstDeclarationBlock, viewModel.IndentFirstDeclarationBlock);
        }

        [TestMethod]
        public void IndentSpacesIsSetInCtor()
        {
            var defaultConfig = GetDefaultConfig();
            var viewModel = new IndenterSettingsViewModel(defaultConfig);

            Assert.AreEqual(defaultConfig.UserSettings.IndenterSettings.IndentSpaces, viewModel.IndentSpaces);
        }
    }
}