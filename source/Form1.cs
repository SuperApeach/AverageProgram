using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AverageProgram
{
    public partial class Form1 : Form
    {
        Calculator cal;

        public Form1()
        {
            InitializeComponent();
            //string path = Application.StartupPath + @"\example2.xlsx";
            //cal = new Calculator(path);
            //cal.Average();
            //PrintDataTable(cal);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string file_path = null;  // openFileDialog : 파일 열기 창 띄우는 컴포넌트
            openFileDialog1.InitialDirectory = Application.StartupPath; // startupPath : 현재프로그램의 경로. winform에서만 사용가능
                                                // initialDirectory : 파일 열기 시 처음 나오는 파일경로(폴더)

            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file_path = openFileDialog1.FileName;
                // .xls  .xlsx 만 받도록 예외처리 해주기
                //label2.Text = file_path.Split('\\')[file_path.Split('\\').Length - 1];

                //FileIO fio = new FileIO();
                //fio.ReadExcel(file_path);

                //////////// test code
                //file_path = Application.StartupPath + @"\example2.xlsx";

                if(file_path.Contains(".xlsx") || file_path.Contains(".xls") || file_path.Contains(".cell"))
                {
                    cal = new Calculator(file_path);
                    PrintDataTable(cal);
                }
                else
                {
                    MessageBox.Show("엑셀파일을 넣어주시기 바랍니다.");
                }
                

            }
        }

        private void PrintDataTable(Calculator cal)
        {
            DataTable table = new DataTable(); // datagridview에 출력하기 위한 자료형식

            // 칼럼 내용 채우는 부분
            table.Columns.Add(new DataColumn(" ", typeof(string))); // 첫번째 칸은 비워두기
            foreach(string item in cal.GetColumn()) // 두번째 칸부터 칼럼 채우기
            {
                table.Columns.Add(new DataColumn(item, typeof(string)));
            }


            // 0번째 칸에는 그룹명, 1번째칸부터 데이터 내용 채워서 테이블에 추가
            // 테스트 필요
            double[,] avg = cal.GetAvg();
            List<string> group = cal.GetGroup();
            int groupCount = group.Count;
            int j = 0;
            foreach(string item in group)
            {   
                string[] temp = new string[1 + avg.Length / groupCount];
                temp[0] = item;
                for(int i = 0;i< avg.Length / groupCount; i++)
                {
                    temp[i + 1] = avg[j, i].ToString();
                }

                table.Rows.Add(temp);

                j++;
            }
            
            

            dataGridView1.DataSource = table;
        }
    }
}
