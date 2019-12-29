using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace Canducci.Excel.UnitTest
{
   [TestClass]
   public class UnitTestHeaderCollection
   {
      [TestMethod]
      public void HeaderCollectionFormatNames()
      {
         HeaderCollection headers0 = new HeaderCollection("n1", "n2", "n3");
         Assert.AreEqual(headers0.Count, 3);
         Assert.AreEqual(headers0[0].Title, "n1");
         Assert.AreEqual(headers0[1].Title, "n2");
         Assert.AreEqual(headers0[2].Title, "n3");

         Assert.AreEqual(headers0[0].Order, 1);
         Assert.AreEqual(headers0[1].Order, 2);
         Assert.AreEqual(headers0[2].Order, 3);
      }

      [TestMethod]
      public void HeaderCollectionTest()
      {
         HeaderCollection headers0 = new HeaderCollection(new Header("t0", 1));
         HeaderCollection headers1 = new HeaderCollection(new Header("t1", 2));
         HeaderCollection headers2 = new HeaderCollection(
            new Header[] {
               new Header("t2", 3),
               new Header("t3", 4)
         });

         HeaderCollection headers3 = new HeaderCollection(new List<Header>(2) {
            new Header("t2", 3),
            new Header("t3", 4)
         });

         HeaderCollection headers4 = new HeaderCollection(10, "p");
         HeaderCollection headers5 = new HeaderCollection(10);

         Assert.AreEqual(headers0.Count, 1);
         Assert.AreEqual(headers0[0].Title, "t0");
         Assert.AreEqual(headers0[0].Order, 1);

         Assert.AreEqual(headers1.Count, 1);
         Assert.AreEqual(headers1[0].Title, "t1");
         Assert.AreEqual(headers1[0].Order, 2);

         Assert.AreEqual(headers2.Count, 2);
         Assert.AreEqual(headers2[0].Title, "t2");
         Assert.AreEqual(headers2[0].Order, 3);
         Assert.AreEqual(headers2[1].Title, "t3");
         Assert.AreEqual(headers2[1].Order, 4);

         Assert.AreEqual(headers3.Count, 2);
         Assert.AreEqual(headers3[0].Title, "t2");
         Assert.AreEqual(headers3[0].Order, 3);
         Assert.AreEqual(headers3[1].Title, "t3");
         Assert.AreEqual(headers3[1].Order, 4);

         Assert.AreEqual(headers4.Count, 10);
         Assert.AreEqual(headers5.Count, 10);
      }

      [TestMethod]
      public void HeaderCollectionFabricTest()
      {         
         HeaderCollection headers1 = HeaderCollection.Create(new Header("t1", 2));
         HeaderCollection headers2 = HeaderCollection.Create(
            new Header[] {
               new Header("t2", 3),
               new Header("t3", 4)
         });

         HeaderCollection headers3 = HeaderCollection.Create(new List<Header>(2) {
            new Header("t2", 3),
            new Header("t3", 4)
         });

         HeaderCollection headers4 = HeaderCollection.Create(10, "p");
         HeaderCollection headers5 = HeaderCollection.Create(10);

         Assert.AreEqual(headers1.Count, 1);
         Assert.AreEqual(headers1[0].Title, "t1");
         Assert.AreEqual(headers1[0].Order, 2);

         Assert.AreEqual(headers2.Count, 2);
         Assert.AreEqual(headers2[0].Title, "t2");
         Assert.AreEqual(headers2[0].Order, 3);
         Assert.AreEqual(headers2[1].Title, "t3");
         Assert.AreEqual(headers2[1].Order, 4);

         Assert.AreEqual(headers3.Count, 2);
         Assert.AreEqual(headers3[0].Title, "t2");
         Assert.AreEqual(headers3[0].Order, 3);
         Assert.AreEqual(headers3[1].Title, "t3");
         Assert.AreEqual(headers3[1].Order, 4);

         Assert.AreEqual(headers4.Count, 10);
         Assert.AreEqual(headers5.Count, 10);
      }
   }
}
