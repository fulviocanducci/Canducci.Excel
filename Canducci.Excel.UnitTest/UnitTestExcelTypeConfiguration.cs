using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Canducci.Excel.UnitTest
{
   [TestClass]
   public class UnitTestExcelTypeConfiguration
   {
      [TestMethod]
      public void ExcelTypeConfigurationTest()
      {
         ExcelTypeConfiguration c0 = new ExcelTypeConfiguration();
         ExcelTypeConfiguration c1 = new ExcelTypeConfiguration(HeaderCollection.Create(new Header("h1", 1)));
         ExcelTypeConfiguration c2 = new ExcelTypeConfiguration("dd/MM/yyyy");
         ExcelTypeConfiguration c3 = new ExcelTypeConfiguration(HeaderCollection.Create(new Header("h1", 1)), "dd/MM/yyyy");

         Assert.AreEqual(c0.DateFormat, "yyyy-MM-dd");
         Assert.AreEqual(c1.DateFormat, "yyyy-MM-dd");
         Assert.AreEqual(c2.DateFormat, "dd/MM/yyyy");
         Assert.AreEqual(c3.DateFormat, "dd/MM/yyyy");

         Assert.AreEqual(c0.DecimalFormat, "#,##0.00");
         Assert.AreEqual(c1.DecimalFormat, "#,##0.00");
         Assert.AreEqual(c2.DecimalFormat, "#,##0.00");
         Assert.AreEqual(c3.DecimalFormat, "#,##0.00");

         Assert.IsNull(c0.Headers);
         Assert.IsNotNull(c1.Headers);
         Assert.IsNull(c2.Headers);
         Assert.IsNotNull(c3.Headers);

         Assert.IsTrue(c1.Headers.Count == 1);
         Assert.IsTrue(c3.Headers.Count == 1);
      }
   }
}
