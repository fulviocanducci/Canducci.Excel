namespace Canducci.Excel.Interfaces
{
    public interface IExcelTypeConfiguration
    {
        string DateFormat { get; set; }
        string DecimalFormat { get; set; }
        HeaderCollection Headers { get; set; }
    }
}