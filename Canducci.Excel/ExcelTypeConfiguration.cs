using Canducci.Excel.Interfaces;
namespace Canducci.Excel
{
   public sealed class ExcelTypeConfiguration : IExcelTypeConfiguration
   {
      private const string MessageDataInvalid = "DateFormat invalid";
      private const string MessageDecimalInvalid = "DecimalFormat invalid";

      public HeaderCollection Headers { get; set; } = null;
      public string DateFormat { get; set; } = "yyyy-MM-dd";
      public string DecimalFormat { get; set; } = "#,##0.00";

      public ExcelTypeConfiguration() { }

      public ExcelTypeConfiguration(string dateFormat, string decimalFormat = "#,##0.00")
      {
         if (string.IsNullOrEmpty(dateFormat))
         {
            throw new System.ArgumentException(MessageDataInvalid, nameof(dateFormat));
         }

         if (string.IsNullOrEmpty(decimalFormat))
         {
            throw new System.ArgumentException(MessageDecimalInvalid, nameof(decimalFormat));
         }

         DateFormat = dateFormat;
         DecimalFormat = decimalFormat;
      }

      public ExcelTypeConfiguration(HeaderCollection headers)
      {
         Headers = headers ?? throw new System.ArgumentNullException(nameof(headers));
      }

      public ExcelTypeConfiguration(HeaderCollection headers, string dateFormat, string decimalFormat = "#,##0.00")
      {
         if (string.IsNullOrEmpty(dateFormat))
         {
            throw new System.ArgumentException(MessageDataInvalid, nameof(dateFormat));
         }

         if (string.IsNullOrEmpty(decimalFormat))
         {
            throw new System.ArgumentException(MessageDecimalInvalid, nameof(decimalFormat));
         }

         Headers = headers;
         DateFormat = dateFormat;
         DecimalFormat = decimalFormat;
      }

      public static IExcelTypeConfiguration Create()
      {
         return new ExcelTypeConfiguration();
      }

      public static IExcelTypeConfiguration Create(HeaderCollection headers)
      {
         return new ExcelTypeConfiguration(headers);
      }

      public static IExcelTypeConfiguration Create(HeaderCollection headers, string dateFormat, string decimalFormat = "#,##0.00")
      {
         return new ExcelTypeConfiguration(headers, dateFormat, decimalFormat);
      }
   }
}
