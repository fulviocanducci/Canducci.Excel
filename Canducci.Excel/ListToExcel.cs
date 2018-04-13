using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Canducci.Excel.Interfaces;
using ClosedXML.Excel;
namespace Canducci.Excel
{
    public class ListToExcel<T> : IListToExcel<T>
    {
        private XLWorkbook _excel;
        private IXLWorksheet _work;
        private IHeaderCollection _headers;
        private string _dateFormat;
        private string _decimalFormat;
        private IEnumerable _items;

#region Load
        protected void Load(IEnumerable items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {            
            _dateFormat = dateFormat;
            _decimalFormat = decimalFormat;
            _headers = headers;
            _items = items;           

            _excel = new XLWorkbook();
            _work = _excel.AddWorksheet("Plan1", 1);

            SetHeaders();
            SetDatas();
            SetAdjustToContents();

        }
#endregion Load

#region Constructors
        public ListToExcel(IEnumerable items, IExcelTypeConfiguration configuration)
        {
            Load(items, configuration.Headers, configuration.DateFormat, configuration.DecimalFormat);
        }
        public ListToExcel(IEnumerable items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            Load(items, headers, dateFormat, decimalFormat);
        }
        public ListToExcel(IEnumerable<T> items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")
        {
            Load(items, headers, dateFormat, decimalFormat);
        }
        public ListToExcel(IQueryable<T> items, IHeaderCollection headers = null, string dateFormat = "dd/MM/yyyy", string decimalFormat = "#,##0.00")            
        {
            Load(items, headers, dateFormat, decimalFormat);
        }
#endregion

#region Configuration        
        private void SetAdjustToContents()
        {
            _work.Columns().AdjustToContents();            
        }
        private void SetTypeDataAndFormatType(ref IXLCell _cell, PropertyInfo _info)
        {
            string _typeString = _info.PropertyType.IsGenericType ?
                _info.PropertyType.GetGenericArguments()[0].Name :
                _info.PropertyType.Name;

            SetTypeDataAndFormatType(ref _cell, _typeString);
        }
        private void SetTypeDataAndFormatType(ref IXLCell _cell, string typeString)
        {                        

            switch (typeString.ToLower())
            {
                case "int":
                case "int16":
                case "int32":
                case "int64":                
                case "uint16":
                case "uint32":
                case "uint64":
                case "short":
                case "ushort":
                case "byte":
                case "sbyte":
                case "bit":
                case "long":
                case "ulong":
                    {
                        _cell.DataType = XLDataType.Number;
                        _cell.Style.NumberFormat.SetFormat("");                        
                        break;
                    }
                case "double":
                case "decimal":
                    {
                        _cell.DataType = XLDataType.Number;
                        _cell.Style.NumberFormat.SetFormat(_decimalFormat);
                        break;
                    }
                case "datetime":
                    {
                        _cell.DataType = XLDataType.DateTime;
                        _cell.Style.DateFormat.SetFormat(_dateFormat);
                        break;
                    }
                case "bool":
                    {
                        _cell.DataType = XLDataType.Boolean;
                        _cell.Style.NumberFormat.SetFormat("");
                        break;
                    }
                case "timespan":
                    {
                        _cell.DataType = XLDataType.TimeSpan;
                        _cell.Style.NumberFormat.SetFormat("");
                        break;
                    }
                default:
                    {
                        _cell.DataType = XLDataType.Text;
                        _cell.Style.NumberFormat.SetFormat("");
                        break;
                    }
            }            
        }
#endregion Configuration

#region SetHeaders
        private void SetHeadersCollection()
        {
            for (int i = 0; i < _headers.Count; i++)
            {
                IXLCell _cellHeader = _work.Cell(1, _headers[i].Order);
                _cellHeader.Value = _headers[i].Title;
                _cellHeader.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                _cellHeader.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            }
        }
        private void SetHeadersDefaultModel()
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            int col = 1;
            foreach(PropertyDescriptor _info in properties)
            {
                IXLCell _cell = _work.Cell(1, col++);
                _cell.Value = _info.Name;                
                _cell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                _cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);                
            }
        }
        private void SetHeaders()
        {
            if (_headers != null && _headers.Count > 0)
            {
                SetHeadersCollection();
            }
            else
            {
                SetHeadersDefaultModel();
            }
        }
#endregion SetHeaders

#region SetDatas
        private void SetDatasInClass(ref int col, ref int row, PropertyInfo[] _infos, dynamic _item)
        {
            col = 1;
            foreach (PropertyInfo _info in _infos)
            {
                IXLCell _cell = _work.Cell(row, col);
                _cell.Value = _info.GetValue(_item, null);
                SetTypeDataAndFormatType(ref _cell, _info);
                col++;
            }
            row++;
        }
        private void SetDatasInArray(ref int col, ref int row, dynamic _item)
        {
            IXLCell _cell = _work.Cell(row, col);
            _cell.Value = _item;
            SetTypeDataAndFormatType(ref _cell, _item.GetType().Name);
            col++;
            if (col > ((Array)_items).Rank)
            {
                col = 1;
                row++;
            }
        }
        private void SetDatasInDictionary(ref int col, ref int row, dynamic _item)
        {
            IXLCell _cell0 = _work.Cell(row, col);
            _cell0.Value = _item.Key;
            SetTypeDataAndFormatType(ref _cell0, _item.Key);
            col++;
            IXLCell _cell1 = _work.Cell(row, col);
            _cell1.Value = _item.Value;
            SetTypeDataAndFormatType(ref _cell1, _item.Value);
            row++;
            col = 1;
        }
        private void SetDatasInPrimitive(ref int col, ref int row, dynamic _item)
        {
            IXLCell _cell = _work.Cell(row++, col);
            _cell.Value = _item;
            SetTypeDataAndFormatType(ref _cell, _item.GetType().Name);
        }
        private void SetDatas()
        {
            int row = 2;
            int col = 1;
            foreach (var _item in _items)
            {
                Type _type = _item.GetType();
                if (_type.IsClass && _type.Name.ToLower().Equals("string") == false)
                {
                    SetDatasInClass(ref col, ref row, _type.GetProperties(), _item);
                }
                else if (_items is Array && ((Array)_items).Rank >= 1)
                {
                    SetDatasInArray(ref col, ref row, _item);
                }
                else if (_items.GetType().GetGenericTypeDefinition() == typeof(Dictionary<,>))
                {
                    SetDatasInDictionary(ref col, ref row, _item);                    
                }
                else
                {
                    SetDatasInPrimitive(ref col, ref row, _item);
                }
            }
        }
#endregion SetDatas

#region SaveAs
        public void SaveAs(string path)
        {
            _excel.SaveAs(path);                                 
        }
        public void SaveAs(System.IO.Stream stream)
        {
            _excel.SaveAs(stream);            
        }
#endregion SaveAs

#region IDisposableSupport
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {                    
                    _headers = null;
                    _items = null;
                }

                _work.Dispose();
                _excel.Dispose();

                disposedValue = true;
            }
        }        
        ~ListToExcel()
        {        
            Dispose(false);
        }        
        public void Dispose()
        {         
            Dispose(true);
            // GC.SuppressFinalize(this);
        }
#endregion IDisposableSupport

    }
}
