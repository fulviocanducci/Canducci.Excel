using Canducci.Excel.Interfaces;
namespace Canducci.Excel
{
    public class Header : IHeader
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

        public static IHeader Create(string title, int order) 
            => new Header(title, order);
    }
}
