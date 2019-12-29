using System;
namespace Canducci.Excel
{
   public class ContentTypeAttribute : Attribute
   {
      public string Value { get; }
      public ContentTypeAttribute(string value)
      {
         if (string.IsNullOrEmpty(value))
         {
            throw new ArgumentException("message", nameof(value));
         }

         Value = value;
      }
   }
}
