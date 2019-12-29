using System.Collections.Generic;

namespace Canducci.Excel
{
   public sealed class HeaderCollection : List<Header>
   {
      public HeaderCollection(params string[] names)
      {
         foreach(string name in names)
         {
            Add(name, (Count + 1));
         }
      }

      public HeaderCollection(IList<Header> headers)
      {
         if (headers is null)
         {
            throw new System.ArgumentNullException(nameof(headers));
         }
         AddRange(headers);
      }

      public HeaderCollection(params Header[] headers)
      {
         if (headers is null)
         {
            throw new System.ArgumentNullException(nameof(headers));
         }

         AddRange(headers);
      }

      public HeaderCollection(int count, string prefix)
      {        
         Add(count, prefix);
      }

      public HeaderCollection(int count)
      {
         Add(count);
      }

      public void Add(string title, int order)
      {
         Add(Header.Create(title, order));
      }

      public static HeaderCollection Create(IList<Header> headers)
      {
         return new HeaderCollection(headers);
      }

      public static HeaderCollection Create(params Header[] headers)
      {
         return new HeaderCollection(headers);
      }

      public static HeaderCollection Create(int count, string prefix)
      {
         return new HeaderCollection(count, prefix);
      }

      public static HeaderCollection Create(int count)
      {
         return new HeaderCollection(count);
      }

      public static HeaderCollection Create(params string[] names)
      {
         return new HeaderCollection(names);
      }

      internal void Add(int count, string prefix)
      {
         if (prefix is null)
         {
            throw new System.ArgumentNullException(nameof(prefix));
         }

         for (int i = 1; i <= count; i++)
         {
            Add(Header.Create($"{prefix}{i}", i));
         }
      }

      internal void Add(int count)
      {
         for (int i = 1; i <= count; i++)
         {
            Add(Header.Create($"{i}", i));
         }
      }
   }
}
