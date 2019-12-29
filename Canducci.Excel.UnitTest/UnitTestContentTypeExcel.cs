using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Canducci.Excel.UnitTest
{
   [TestClass]
   public class UnitTestContentTypeExcel
   {
      [TestMethod]
      public void ContentTypeExcelTest()
      {
         Assert.AreEqual("application/vnd.ms-excel", ContentTypeExcel.Xls.ToValue());
         Assert.AreEqual("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ContentTypeExcel.Xlsx.ToValue());
      }
   }
}
