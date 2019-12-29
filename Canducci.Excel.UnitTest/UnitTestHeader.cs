using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Canducci.Excel.UnitTest
{
   [TestClass]
   public class UnitTestHeader
   {
      [TestMethod]
      public void HeaderTest()
      {
         Header header0 = new Header("t0", 1);
         Header header1 = Header.Create("t1", 2);

         Assert.AreEqual(header0.Title, "t0");
         Assert.AreEqual(header1.Title, "t1");

         Assert.AreEqual(header0.Order, 1);
         Assert.AreEqual(header1.Order, 2);
      }

      [TestMethod]
      [ExpectedException(typeof(System.ArgumentException))]
      public void HeaderExceptionTest()
      {
         _ = new Header("", 1);
         _ = Header.Create("", 2);
      }
   }
}
