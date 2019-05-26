using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace AverageProgram
{
    class FileIO
    {
        Excel.Application excelApp = null;
        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        public object[,] ReadExcel(string path)
        {
            //List<string> inputdata = new List<string>(); // 일반화 컬렉션 : 타입 지정해서 박싱,언박싱 X, 성능 더 좋음
                                                         // 그냥 컬렉션 : ArrayList
            object[,] data;

            try
            {
                //string path;

                excelApp = new Excel.Application();
                wb = excelApp.Workbooks.Open(path);
                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

                Excel.Range last = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range rng = ws.get_Range("A1", last);

                data = rng.Value;
                // data.GetLength(0) 은 행의 갯수(row의 갯수), data.GetLenght(1)은 열의 갯수(column의 갯수)
                // 즉, GetLength(x)  x에 들어가는 숫자는 차원.  2차원배열이면 0은 앞에 차원, 1은 뒤에 차원
                //int a = data.GetLength(0);
                //int b = data.GetLength(1);

                // 우선 한줄로 받는다.  null을 제외하고 모두 입력받을때까지 받아준다.

                //data 형식 : 1부터 시작 (0이 아님) [1,1] ~ [1,6]  ~~~  [n,1]~[n,6]

                /*
                //////////////////// 보류
            
                for(int i = 1; i <= data.GetLength(0); i++)
                {
                    for (int j = 1; j <= data.GetLength(1); j++)
                    {
                        if(data[i,j] != null)
                        {
                            inputdata.Add(data[i, j].ToString());
                        }
                    }
                }
                */
                // 그냥 2차원 object를 통째로 보내고 가서 1차원배열로 변환 없이 데이터 가공?
                // 그게 더 편리

                wb.Close(true);
                excelApp.Quit();
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

            return data;
        }

        public void WriteExcel()
        {

        }
    }
}
