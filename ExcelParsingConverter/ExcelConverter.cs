using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelParsingConverter
{
    class ExcelConverter
    {
        // 시트 번호(1번부터) , column 번호(1번부터), column 데이터들
        Dictionary<int, Dictionary<int, List<string>>> _excelData;
        //시트 이름, 인덱스번호, 컬럼명, 셀 데이터
        Dictionary<string, Dictionary<string, Dictionary<string, string>>> _excelConvert;

        /// <summary>
        /// 지정 객체를 제거(해지)하는 함수.
        /// </summary>
        /// <param name="obj">제거할 객체</param>
        void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// 해당하는 워크 시트의 내용의 실질 컬럼 수를 받는 함수
        /// </summary>
        /// <param name="oSheet">컬럼수를 알고 싶은 시트</param>
        /// <param name="oRng">목표가 되는 컬럼</param>
        /// <returns>컬럼의 개수</returns>
        int ExcelFileColumnCount(Excel.Worksheet oSheet, Excel.Range oRng)
        {
            int colCount = oRng.Count;

            //20은 예시 값이고 이후 확인하기.
            for (int i = 1; i <= 20; i++)
            {
                Excel.Range cell = (Excel.Range)oSheet.Cells[1, i];

                if (cell.Value == null)
                {
                    ReleaseExcelObject(cell);
                    Console.WriteLine(oSheet.Name.ToString() + " Sheet에 비어있는 셀이 존재합니다.");
                    colCount = i - 1;
                    break;
                }
            }

            return colCount;
        }

        /// <summary>
        /// 엑셀이 지정한 ColumnName을 구성해서 List로 반환하는 함수
        /// </summary>
        /// <param name="length">실질 데이터가 들어가 있는 컬럼의 수</param>
        /// <returns>구성된 엑셀 컬럼이름의 List</returns>
        List<string> ExcelFileColumnsName(int length)
        {
            List<string> columnList = new List<string>();

            int baseNum = 26;                               //알파벳 수

            for (int i = 0; i < length; i++)
            {
                if (i / baseNum == 0)
                {
                    columnList.Add(Convert.ToString((char)(65 + i)));
                }
                else
                {
                    string tempData = Convert.ToString((char)(64 + (i / baseNum)));
                    tempData += Convert.ToString((char)(65 + (i % baseNum)));
                    columnList.Add(Convert.ToString(tempData));
                }
            }
            return columnList;
        }


        bool ExcelFileLoad(in string path)
        {
            object misValue = System.Reflection.Missing.Value;

            Excel.Application oXL = new Excel.Application();
            Excel.Workbooks oWBooks = oXL.Workbooks;
            Excel.Workbook oWB = oWBooks.Open(path, misValue, misValue, misValue, misValue, misValue, misValue,
                                              misValue, misValue, misValue, misValue, misValue, misValue, misValue, misValue);
            Excel.Sheets oSheets = oWB.Worksheets;

            try
            {
                // oSheets 안에 있는 정보를 Dictionary로 저장.
                ExcelFileSaveToDictionary(oSheets);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }

            oXL.Visible = false;
            oXL.UserControl = true;
            oXL.DisplayAlerts = false;
            oXL.Quit();

            ReleaseExcelObject(oSheets);
            ReleaseExcelObject(oWB);
            ReleaseExcelObject(oWBooks);
            ReleaseExcelObject(oXL);

            return true;
        }


        void ExcelFileSaveToDictionary(Excel.Sheets oSheets)
        {
            for (int i = 1; i <= oSheets.Count; i++)
            {
                List<string> columns;

                Excel.Worksheet oSheet = (Excel.Worksheet)oSheets.get_Item(i);
                Excel.Range oRng = oSheet.get_Range("A1").SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

                int colCount = ExcelFileColumnCount(oSheet, oRng);
                columns = ExcelFileColumnsName(colCount);

                Dictionary<int, List<string>> dicSheet = new Dictionary<int, List<string>>();
                for (int j = 1; j <= colCount; j++)
                {
                    List<string> listData = new List<string>();
                    int count = 0;

                    Excel.Range collCell = (Excel.Range)oSheet.Columns[j];
                    Excel.Range range = oSheet.get_Range(columns[j - 1] + "1", collCell);

                    foreach (object item in range.Value)
                    {
                        if (count < oRng.Row)
                        {
                            count++;
                            if (item == null)
                                listData.Add("");
                            else
                                listData.Add(item.ToString());
                        }
                        else break;
                    }
                    dicSheet.Add(j, listData);
                    ReleaseExcelObject(range);
                    ReleaseExcelObject(collCell);
                }
                _excelData.Add(i, dicSheet);
            }
        }

        public void InitConvert(in string fileName)
        {
            string fullName = Directory.GetCurrentDirectory() + "\\" + fileName;

            _excelData = new Dictionary<int, Dictionary<int, List<string>>>();

            // 액셀로드하여 데이터 저장.
            if (!ExcelFileLoad(fullName))
            {
                Console.WriteLine("파일 로드에 실패했습니다");
            }
            else
            {
                Console.WriteLine("엑셀데이터 파싱 종료");
            }
        }

        public void ConversionData()
        {
            _excelConvert = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();
            foreach (var sheet in _excelData.Values)
            {
                Dictionary<string, Dictionary<string, string>> dicSheet = new Dictionary<string, Dictionary<string, string>>();
                _excelConvert.Add(sheet.ToString(), dicSheet);
                foreach (var column in sheet.Values)
                {
                    Dictionary<string, string> dicCol = new Dictionary<string, string>();

                    for (int i = 1; i < column.Count; i++)
                    {
                        dicCol.Add(column[0], column[i]);
                    }
                    dicSheet.Add(sheet[1].ToString(), dicCol);
                }
            }
        }
        public void ShowOriginDictionary()
        {
            foreach (var sheet in _excelConvert)
            {
                Console.WriteLine(sheet.Key.ToString() + "===============\n");
                //시트 별로 column 별 한 줄씩 나열.
                foreach (var column in sheet.Value)
                {
                    Console.WriteLine(column.Key.ToString() + "===============");
                    foreach (var item in column.Value)
                    {
                        Console.Write("<" + item.Key + " : " + item.Value + ">\t");
                    }
                }
                Console.WriteLine();
            }
        }
        public void ShowConvertDictionary()
        {
            //시트 별로 
            foreach (var sheet in _excelData.Values)
            {
                //시트 별로 column 별 한 줄씩 나열.
                foreach (var column in sheet.Values)
                {
                    foreach (var item in column)
                    {
                        Console.Write(item);
                        Console.Write("   ");
                    }
                }

                Console.WriteLine();
            }
        }
    }
}
