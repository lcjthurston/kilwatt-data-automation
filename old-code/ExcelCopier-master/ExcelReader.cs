using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace TesteExcelInterface
{
    public class ExcelReader : IExcelManager
    {
        public string? Path { get; set; }
        public ExcelReader(string? path)
        {
            Path = path;
            this.Read();
        }

        public void Read()
        {
            throw new NotImplementedException();
        }

        public void Save()
        {
            throw new NotImplementedException();
        }

        public void Save(string path)
        {
            throw new NotImplementedException();
        }
    }
}
