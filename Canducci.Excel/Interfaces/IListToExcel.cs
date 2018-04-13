namespace Canducci.Excel.Interfaces
{
    public interface IListToExcel<T>: System.IDisposable
    {
        void SaveAs(string path);
        void SaveAs(System.IO.Stream stream);
    }
}