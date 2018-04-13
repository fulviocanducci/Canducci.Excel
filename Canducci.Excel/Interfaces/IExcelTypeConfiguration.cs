namespace Canducci.Excel.Interfaces
{
    public interface IExcelTypeConfiguration
    {
        string DateFormat { get;  }
        string DecimalFormat { get;  }
        IHeaderCollection Headers { get; }
    }
}