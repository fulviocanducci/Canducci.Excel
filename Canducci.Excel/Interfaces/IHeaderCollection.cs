using System.Collections.Generic;
namespace Canducci.Excel.Interfaces
{
    public interface IHeaderCollection: IList<IHeader>            
    {
        void Add(int count, string prefix);
        void Add(int count);
        void Add(string title, int order);        
    }
}