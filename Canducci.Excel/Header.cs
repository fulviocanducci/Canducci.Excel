namespace Canducci.Excel
{
   public class Header
   {
      public string Title { get; }
      public int Order { get; }

      public Header(string title, int order)
      {
         if (string.IsNullOrEmpty(title))
         {
            throw new System.ArgumentException("Title not empty", nameof(title));
         }

         Title = title;
         Order = order;
      }

      public static Header Create(string title, int order)
      {
         return new Header(title, order);
      }
   }
}
