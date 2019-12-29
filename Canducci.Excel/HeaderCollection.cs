using Canducci.Excel.Interfaces;
using System.Collections.Generic;
using System.Linq;

namespace Canducci.Excel
{
   public sealed class HeaderCollection : List<IHeader>, IHeaderCollection
   {
      public HeaderCollection() { }

      public HeaderCollection(IList<IHeader> headers)
      {
         if (headers == null)
         {
            throw new System.ArgumentNullException(nameof(headers));
         }

         AddRange(headers);
      }

      public HeaderCollection(params IHeader[] headers)
          : this(headers.ToList())
      {
      }

      public void Add(string title, int order)
          => Add(Header.Create(title, order));

      public void Add(int count, string prefix)
      {
         for (int i = 1; i <= count; i++)
         {
            Add(Header.Create($"{prefix}{i}", i));
         }
      }

      public void Add(int count)
      {
         for (int i = 1; i <= count; i++)
         {
            Add(Header.Create($"{i}", i));
         }
      }

      public static IHeaderCollection Create()
          => new HeaderCollection();

      public static IHeaderCollection Create(IList<IHeader> headers)
          => new HeaderCollection(headers);

   }
}
