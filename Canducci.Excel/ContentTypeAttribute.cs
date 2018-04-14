using System;
namespace Canducci.Excel
{    
    public class ContentTypeAttribute: Attribute 
    {
        public string Value { get; }
        public ContentTypeAttribute(string value)
        {
            Value = value;
        }
    }
}
