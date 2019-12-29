using System.IO;
namespace Canducci.Excel.Internals
{
   internal static class FabricMemoryStream
    {
        public static MemoryStream Create()
        {
            return new MemoryStream();
        }
    }
}
