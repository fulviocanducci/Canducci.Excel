using Canducci.Excel.Interfaces;

namespace Canducci.Excel
{
   public sealed class HeaderCreator
   {
      private IHeaderCollection HeaderCollection { get; set; }

      public HeaderCreator()
      {
         HeaderCollection = new HeaderCollection();
      }

      public HeaderCreator Add(string title, int order)
      {
         HeaderCollection.Add(title, order);
         return this;
      }

      public HeaderCreator Add(IHeader header)
      {
         HeaderCollection.Add(header);
         return this;
      }

      public IHeaderCollection Render()
      {
         return HeaderCollection;
      }

      public static HeaderCreator Create()
          => new HeaderCreator();
   }
}
