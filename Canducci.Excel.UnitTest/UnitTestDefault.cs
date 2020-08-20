using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
namespace Canducci.Excel.UnitTest
{
   [TestClass]
   public class UnitTestDefault
   {
      [TestMethod]
      public void TestListToExcel()
      {
         IList<People> peoples = new List<People>();
         peoples.Add(new People
         {
            Id = 1,
            Name = "Name 1"
         });
         peoples.Add(new People
         {
            Id = 2,
            Name = "Name 2"
         });

         byte[] peopleExcelArrayBytes = peoples.ToExcelByte();
         bool save = peoples.ToExcelSaveAs($"./{System.Guid.NewGuid().ToString()}.xlsx");

         Assert.IsInstanceOfType(peopleExcelArrayBytes.GetType(), typeof(byte[]).GetType());
         Assert.IsTrue(save);
      }

      [TestMethod]
      public void TestListToExcelConfiguration()
      {
            IList<People> peoples = new List<People>
            {
                new People
                {
                    Id = 1,
                    Name = "Name 1"
                },
                new People
                {
                    Id = 2,
                    Name = "Name 2"
                }
            };

            ExcelTypeConfiguration config = new ExcelTypeConfiguration();         
         byte[] peopleExcelArrayBytes = peoples.ToExcelByte();
         bool save = peoples.ToExcelSaveAs($"./{System.Guid.NewGuid().ToString()}.xlsx", config);

         Assert.IsInstanceOfType(peopleExcelArrayBytes.GetType(), typeof(byte[]).GetType());
         Assert.IsTrue(save);
      }
   }
}
