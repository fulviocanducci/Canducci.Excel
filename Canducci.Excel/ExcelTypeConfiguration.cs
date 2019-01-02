using Canducci.Excel.Interfaces;
namespace Canducci.Excel
{
    public sealed class ExcelTypeConfiguration : IExcelTypeConfiguration
    {
        private const string Message_DataInvalid = "DateFormat invalid";
        private const string Message_DecimalInvalid = "DecimalFormat invalid";

        public IHeaderCollection Headers { get; set; } = null;
        public string DateFormat { get; set; } = "dd/MM/yyyy";
        public string DecimalFormat { get; set; } = "#,##0.00"; 
        
        public ExcelTypeConfiguration() { }

        public ExcelTypeConfiguration(string dateFormat, string decimalFormat = "#,##0.00")
        {
            if (string.IsNullOrEmpty(dateFormat))
            {
                throw new System.ArgumentException(Message_DataInvalid, nameof(dateFormat));
            }

            if (string.IsNullOrEmpty(decimalFormat))
            {
                throw new System.ArgumentException(Message_DecimalInvalid, nameof(decimalFormat));
            }

            DateFormat = dateFormat;
            DecimalFormat = decimalFormat;
        }

        public ExcelTypeConfiguration(IHeaderCollection headers) 
            => Headers = headers ?? throw new System.ArgumentNullException(nameof(headers));

        public ExcelTypeConfiguration(IHeaderCollection headers, string dateFormat, string decimalFormat = "#,##0.00")
        {
            if (string.IsNullOrEmpty(dateFormat))
            {
                throw new System.ArgumentException(Message_DataInvalid, nameof(dateFormat));
            }

            if (string.IsNullOrEmpty(decimalFormat))
            {
                throw new System.ArgumentException(Message_DecimalInvalid, nameof(decimalFormat));
            }

            Headers = headers;
            DateFormat = dateFormat;
            DecimalFormat = decimalFormat;
        }

        public static IExcelTypeConfiguration Create() 
            => new ExcelTypeConfiguration();

        public static IExcelTypeConfiguration Create(IHeaderCollection headers) 
            => new ExcelTypeConfiguration(headers);

        public static IExcelTypeConfiguration Create(IHeaderCollection headers, string dateFormat, string decimalFormat = "#,##0.00") 
            => new ExcelTypeConfiguration(headers, dateFormat, decimalFormat);
    }
}
