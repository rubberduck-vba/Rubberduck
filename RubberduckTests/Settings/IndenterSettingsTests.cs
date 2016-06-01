using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;
using Rubberduck.UI;
using Rubberduck.UI.Settings;

namespace RubberduckTests.Settings
{
    [TestClass]
    public class IndenterSettingsTests
    {
        #region Defaults
        private const bool DefaultIndentEntireProcedureBody = true;
        private const bool DefaultIndentFirstCommentBlock = true;
        private const bool DefaultIndentFirstDeclarationBlock = true;
        private const bool DefaultAlignCommentsWithCode = true;
        private const bool DefaultAlignContinuations = true;
        private const bool DefaultIgnoreOperatorsInContinuations = true;
        private const bool DefaultIndentCase = false;
        private const bool DefaultForceDebugStatementsInColumn1 = false;
        private const bool DefaultForceCompilerDirectivesInColumn1 = false;
        private const bool DefaultIndentCompilerDirectives = true;
        private const bool DefaultAlignDims = false;
        private const int DefaultAlignDimColumn = 15;
        private const bool DefaultEnableUndo = true;
        private const EndOfLineCommentStyle DefaultEndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
        private const int DefaultEndOfLineCommentColumnSpaceAlignment = 50;
        private const int DefaultIndentSpaces = 4;
        #endregion

        #region Nondefaults
        private const bool NondefaultIndentEntireProcedureBody = false;
        private const bool NondefaultIndentFirstCommentBlock = false;
        private const bool NondefaultIndentFirstDeclarationBlock = false;
        private const bool NondefaultAlignCommentsWithCode = false;
        private const bool NondefaultAlignContinuations = false;
        private const bool NondefaultIgnoreOperatorsInContinuations = false;
        private const bool NondefaultIndentCase = true;
        private const bool NondefaultForceDebugStatementsInColumn1 = true;
        private const bool NondefaultForceCompilerDirectivesInColumn1 = true;
        private const bool NondefaultIndentCompilerDirectives = false;
        private const bool NondefaultAlignDims = true;
        private const int NondefaultAlignDimColumn = 16;
        private const bool NondefaultEnableUndo = false;
        private const EndOfLineCommentStyle NondefaultEndOfLineCommentStyle = EndOfLineCommentStyle.Absolute;
        private const int NondefaultEndOfLineCommentColumnSpaceAlignment = 60;
        private const int NondefaultIndentSpaces = 2;
        #endregion

        public static Rubberduck.SmartIndenter.IndenterSettings GetMockIndenterSettings(bool nondefault = false)
        {
            var output = new Mock<Rubberduck.SmartIndenter.IndenterSettings>();

            output.SetupProperty(s => s.IndentEntireProcedureBody);
            output.SetupProperty(s => s.IndentFirstCommentBlock);
            output.SetupProperty(s => s.IndentFirstDeclarationBlock);
            output.SetupProperty(s => s.AlignCommentsWithCode);
            output.SetupProperty(s => s.AlignContinuations);
            output.SetupProperty(s => s.IgnoreOperatorsInContinuations);
            output.SetupProperty(s => s.IndentCase);
            output.SetupProperty(s => s.ForceDebugStatementsInColumn1);
            output.SetupProperty(s => s.ForceCompilerDirectivesInColumn1);
            output.SetupProperty(s => s.IndentCompilerDirectives);
            output.SetupProperty(s => s.AlignDims);
            output.SetupProperty(s => s.AlignDimColumn);
            output.SetupProperty(s => s.EnableUndo);
            output.SetupProperty(s => s.EndOfLineCommentStyle);
            output.SetupProperty(s => s.EndOfLineCommentColumnSpaceAlignment);
            output.SetupProperty(s => s.IndentSpaces);

            output.Object.IndentEntireProcedureBody = nondefault ? NondefaultIndentEntireProcedureBody : DefaultIndentEntireProcedureBody;
            output.Object.IndentFirstCommentBlock = nondefault ? NondefaultIndentFirstCommentBlock : DefaultIndentFirstCommentBlock;
            output.Object.IndentFirstDeclarationBlock = nondefault ? NondefaultIndentFirstDeclarationBlock : DefaultIndentFirstDeclarationBlock;
            output.Object.AlignCommentsWithCode = nondefault ? NondefaultAlignCommentsWithCode : DefaultAlignCommentsWithCode;
            output.Object.AlignContinuations = nondefault ? NondefaultAlignContinuations : DefaultAlignContinuations;
            output.Object.IgnoreOperatorsInContinuations = nondefault ? NondefaultIgnoreOperatorsInContinuations : DefaultIgnoreOperatorsInContinuations;
            output.Object.IndentCase = nondefault ? NondefaultIndentCase : DefaultIndentCase;
            output.Object.ForceDebugStatementsInColumn1 = nondefault ? NondefaultForceDebugStatementsInColumn1 : DefaultForceDebugStatementsInColumn1;
            output.Object.ForceCompilerDirectivesInColumn1 = nondefault ? NondefaultForceCompilerDirectivesInColumn1 : DefaultForceCompilerDirectivesInColumn1;
            output.Object.IndentCompilerDirectives = nondefault ? NondefaultIndentCompilerDirectives : DefaultIndentCompilerDirectives;
            output.Object.AlignDims = nondefault ? NondefaultAlignDims : DefaultAlignDims;
            output.Object.AlignDimColumn = nondefault ? NondefaultAlignDimColumn : DefaultAlignDimColumn;
            output.Object.EnableUndo = nondefault ? NondefaultEnableUndo : DefaultEnableUndo;
            output.Object.EndOfLineCommentStyle = nondefault ? NondefaultEndOfLineCommentStyle : DefaultEndOfLineCommentStyle;
            output.Object.EndOfLineCommentColumnSpaceAlignment = nondefault ? NondefaultEndOfLineCommentColumnSpaceAlignment : DefaultEndOfLineCommentColumnSpaceAlignment;
            output.Object.IndentSpaces = nondefault ? NondefaultIndentSpaces : DefaultIndentSpaces;

            return output.Object;
        }

        private Configuration GetDefaultConfig()
        {
            //var indenterSettings = GetTestIndenterSettings();
            //{
            //    indenterSettings.IndentEntireProcedureBody = true,
            //    indenterSettings.IndentFirstCommentBlock = true,
            //    indenterSettings.IndentFirstDeclarationBlock = true,
            //    indenterSettings.AlignCommentsWithCode = true,
            //    indenterSettings.AlignContinuations = true,
            //    IgnoreOperatorsInContinuations = true,
            //    IndentCase = false,
            //    ForceDebugStatementsInColumn1 = false,
            //    ForceCompilerDirectivesInColumn1 = false,
            //    IndentCompilerDirectives = true,
            //    AlignDims = false,
            //    AlignDimColumn = 15,
            //    EnableUndo = true,
            //    EndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn,
            //    EndOfLineCommentColumnSpaceAlignment = 50,
            //    IndentSpaces = 4
            //};

            var userSettings = new UserSettings(null, null, null, null, null, GetMockIndenterSettings());
            return new Configuration(userSettings);
        }

        private Configuration GetNondefaultConfig()
        {
            //var indenterSettings = new Rubberduck.SmartIndenter.IndenterSettings
            //{
            //    IndentEntireProcedureBody = false,
            //    IndentFirstCommentBlock = false,
            //    IndentFirstDeclarationBlock = false,
            //    AlignCommentsWithCode = false,
            //    AlignContinuations = false,
            //    IgnoreOperatorsInContinuations = false,
            //    IndentCase = true,
            //    ForceDebugStatementsInColumn1 = true,
            //    ForceCompilerDirectivesInColumn1 = true,
            //    IndentCompilerDirectives = false,
            //    AlignDims = true,
            //    AlignDimColumn = 16,
            //    EnableUndo = false,
            //    EndOfLineCommentStyle = EndOfLineCommentStyle.Absolute,
            //    EndOfLineCommentColumnSpaceAlignment = 60,
            //    IndentSpaces = 2
            //};

            var userSettings = new UserSettings(null, null, null, null, null, GetMockIndenterSettings(true));
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
