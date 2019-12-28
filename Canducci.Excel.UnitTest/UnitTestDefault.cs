using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
namespace Canducci.Excel.UnitTest
{
   [TestClass]
   public class UnitTestDefault
   {
      [TestMethod]
      public void TestMethod1()
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
   }
}
