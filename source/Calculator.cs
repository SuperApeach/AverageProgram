using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AverageProgram
{
    class Calculator
    {
        FileIO fio = new FileIO();
        object[,] rawData;
        //string[] column = new string[5]; // 칼럼 갯수도 바뀔수 있음을 주의. 이거 다시 바꿔야됨
        List<string> column = new List<string>();
        List<string> group = new List<string>();
        List<double[]> data = new List<double[]>();
        double[,] avg;

        public Calculator(string path)
        {
            rawData = fio.ReadExcel(path);
            ProcessData();
            Average();
        }

        private void ProcessData()
        {
            for(int i = 2; i <= rawData.GetLength(1); i++)
            {
                //column[i - 2] = rawData[1, i].ToString();
                if(rawData[1,i] != null)
                {
                    column.Add(rawData[1, i].ToString());
                }
            }

            for(int i = 2; rawData[i, 1] != null; i++)
            {
                group.Add(rawData[i, 1].ToString());
            }
            // 액셀에서 0.5% 1% 이렇게 입력된 경우, 실제 데이터는 0.005, 0.01 로 들어오니 표기에 주의.

            for(int i = 2; i <= rawData.GetLength(0); i++)
            {
                double[] row = new double[5];

                if(rawData[i,2] != null)
                {
                    for (int j = 2; j <= column.Count+1; j++) // get length(1) 로 하면 안됨. column의 갯수 + 1로 세어주기
                    {
                        row[j - 2] = Convert.ToDouble(rawData[i, j]); //rawData[i, j].ToString();
                    }

                    data.Add(row);
                }
            }
        }

        public void Average()
        {
            //double[] a = data.ElementAt(1);
            avg = new double[group.Count,column.Count]; // 평균 역시 2차원 테이블임.
             // 가로로는 칼럼(설문문항)갯수, 세로로는 그룹(실험군 및 대조군) 갯수
            
            for(int i = 0; i < data.Count; i++) // avg 테이블에 모두 위치에 맞춰 더하기
            {
                double[] row = data.ElementAt(i);

                for(int j = 0; j < column.Count; j++)
                {
                    avg[i%group.Count, j] += row[j];
                }
            }

            for(int i = 0; i < group.Count; i++) // 설문자 수 만큼 나누기 (평균)
            {
                for(int j = 0; j < column.Count; j++)
                {
                    avg[i, j] = avg[i, j] / (data.Count / group.Count);
                    avg[i,j] = Math.Round(avg[i, j] * 100) / 100;
                }
            }

            
        }
        
        public double[,] GetAvg()
        {
            return avg;
        }

        public List<string> GetColumn()
        {
            return column;
        }

        public List<string> GetGroup()
        {
            return group;
        }
    }
}
