using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParsingConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelConverter ec = new ExcelConverter();
            //ec.InitConvert("GameTableDoc.xlsx");
            //ec.ShowOriginDictionary();
            //ec.ConversionData();
            ec.GetJsonFile("MonsterTable");
            //ec.ShowConvertDictionary();
            //ec.SaveTextFile('|');
            //ec.SetJsonFile();
        }
    }
}
