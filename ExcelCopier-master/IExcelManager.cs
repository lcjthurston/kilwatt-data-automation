using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesteExcelInterface
{
    internal interface IExcelManager
    {
        string Path { get; set; }
        void Read();
        void Save();
        void Save(string path);
    }
}
