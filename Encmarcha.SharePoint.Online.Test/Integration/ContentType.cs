using Encmarcha.SharePoint.Online.Test.Utils;
using Enmarcha.SharePoint.Abstract.Interfaces.Artefacts;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace Encmarcha.SharePoint.Online.Test.Integration
{
    [TestClass]
    public class ContentType
    {
        public ClientContext ClientContext;        
        [TestInitialize]
        public void Init()
        {
            ClientContext = ContextSharePoint.CreateClientContext();

        }
        [TestMethod]
        public void CreateContentTypeSuccess()
        {
            var mockLogger = new Mock<ILog>();
            var logger = mockLogger.Object;
            var contentType = new Enmarcha.SharePoint.Online.Entities.Artefacts.ContentType(ClientContext, logger,
                "ContentTypeTest", "ENMARCHA", "0x0101");

            var result = contentType.Create(string.Empty);
            Assert.IsTrue(result);
            contentType.Delete();
        }
        [TestMethod]
        public void CreateContentTypeFail()
        {
            var mockLogger = new Mock<ILog>();
            var logger = mockLogger.Object;
            var contentType = new Enmarcha.SharePoint.Online.Entities.Artefacts.ContentType(ClientContext, logger,
                "ContentTypeTestFail", "ENMARCHA", "");

            var result = contentType.Create(string.Empty);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void DeleteContentTypeSucess()
        {
            var mockLogger = new Mock<ILog>();
            var logger = mockLogger.Object;
            var contentType = new Enmarcha.SharePoint.Online.Entities.Artefacts.ContentType(ClientContext, logger,
             "ContentTypeTest", "ENMARCHA", "0x0101");
            var result = contentType.Exist();
            if (!result)
            {
                contentType.Create(string.Empty);
            }
            result= contentType.Delete();
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void AddColumnAtContentTypeSuccess()
        {
            var mockLogger = new Mock<ILog>();
            var logger = mockLogger.Object;
            var contentType = new Enmarcha.SharePoint.Online.Entities.Artefacts.ContentType(ClientContext, logger,
             "ContentTypeTest", "ENMARCHA", "0x0101");
            var result = contentType.Exist();
            if (!result)
            {
                contentType.Create(string.Empty);
            }
            Assert.IsTrue(contentType.AddColumn("Categories"));
            contentType.Delete();
        }

        [TestMethod]
        public void AddColumnAtContentTypeFail()
        {
            var mockLogger = new Mock<ILog>();
            var logger = mockLogger.Object;
            var contentType = new Enmarcha.SharePoint.Online.Entities.Artefacts.ContentType(ClientContext, logger,
             "ContentTypeTest", "ENMARCHA", "0x0101");
            var result = contentType.Exist();
            if (!result)
            {
                contentType.Create(string.Empty);
            }
            Assert.IsFalse(contentType.AddColumn("CategoriesNotExist"));
            contentType.Delete();
        }

        [TestMethod]
        public void RemoveColumnAtContentTypeSuccess()
        {
            var mockLogger = new Mock<ILog>();
            var logger = mockLogger.Object;
            var contentType = new Enmarcha.SharePoint.Online.Entities.Artefacts.ContentType(ClientContext, logger,
             "ContentTypeTest", "ENMARCHA", "0x0101");
            var result = contentType.Exist();
            if (!result)
            {
                contentType.Create(string.Empty);
            }
            contentType.AddColumn("Categories");
            Assert.IsTrue(contentType.RemoveColumn("Categories"));
            contentType.Delete();
        }
    }
}
