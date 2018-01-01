using System;
using ApplyConfigTransformInPlace.VSIX;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ApplyConfigTransformInPlace.Tests
{
    [TestClass]
    public class IsSupportedTranformTests
    {
        [TestMethod]
        public void Web_Config_is_not_supported_transform()
        {
            NegativeAssert("web.config");
        }

        [TestMethod]
        public void Validate_transform_pattern()
        {
            PositiveAssert("web.anything.config","web.");
            PositiveAssert("saml.anything.config","saml.");
            PositiveAssert("app.anything.config","app.");
            PositiveAssert("app.Dev_1.config","app.");
            PositiveAssert("app.Dev-1.config","app.");
            PositiveAssert("app.123.config", "app.");
        }

        [TestMethod]
        public void IsSupportedTransform_is_case_insensative()
        {
            NegativeAssert("web.config");
            NegativeAssert("Web.Config");
        }

        private void PositiveAssert(string transformFileName, string expectedDestinationPrefix)
        {
            string actualPrefix;
            var result = ApplyConfigTransformInPlaceLogic.IsSupportedTransform(transformFileName, out actualPrefix);
            Assert.IsTrue(result, $"The transform file name {transformFileName} was not supported and should have been");
            Assert.AreEqual(expectedDestinationPrefix, actualPrefix);
        }

        private void NegativeAssert(string transformFileName)
        {
            string actualPrefix;
            var result = ApplyConfigTransformInPlaceLogic.IsSupportedTransform(transformFileName, out actualPrefix);
            Assert.IsFalse(result, $"The transform file name {transformFileName} was supported and should not have been");
            Assert.AreEqual(string.Empty, actualPrefix);
        }
    }
}
