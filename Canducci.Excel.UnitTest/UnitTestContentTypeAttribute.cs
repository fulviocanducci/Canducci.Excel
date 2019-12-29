using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Canducci.Excel.UnitTest
{
   [TestClass]
   public class UnitTestContentTypeAttribute
   {
      [TestMethod]
      public void ContentTypeAttributeTest ()
      {
         ContentTypeAttribute ctXlsx = new ContentTypeAttribute("xlsx");
         ContentTypeAttribute ctXls = new ContentTypeAttribute("xls");

         Assert.AreEqual(ctXls.Value, "xls");
         Assert.AreEqual(ctXlsx.Value, "xlsx");
      }

      [TestMethod]
      [ExpectedException(typeof(System.ArgumentException))]
      public void ContentTypeAttributeException()
      {
         ContentTypeAttribute ab = new ContentTypeAttribute("");
         ContentTypeAttribute cd = new ContentTypeAttribute(null);
      }
   }
}
