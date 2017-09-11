using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using MathNet.Numerics.LinearAlgebra;
using Application = Microsoft.Office.Interop.Excel.Application;


namespace Non_Linear_Porabalistic_Chain_WinForm
{
    public partial class Form1 : Form
    {
        private double[,] initialData;
        private int[] arrYears;
        private string[] arrCountry;
        private int col;
        private int row;
        private List<List<double>> result;

        public Form1()
        {
            InitializeComponent();
            col = 0;
            row = 0;
            result = new List<List<double>>();
            periodToolStripComboBox.SelectedIndex = 0;

        }

        private void LoadFromFile(String fileName)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            Excel.Range range = sheet.UsedRange;

            row = range.Rows.Count - 1;
            col = range.Columns.Count - 1;

            arrCountry = new string[col];
            arrYears = new int[row];
            this.initialData = new double[row, col];



            //country
            for (int i = 1; i <= col; i++)
            {
                arrCountry[i - 1] = (range.Cells[1, i + 1] as Excel.Range).Value;
            }
            for (int i = 1; i <= row; i++)
            {
                arrYears[i - 1] = (int)(range.Cells[i + 1, 1] as Excel.Range).Value2;
            }

            for (int i = 2; i <= row + 1; i++)
            {
                for (int j = 2; j <= col + 1; j++)
                {
                    this.initialData[i - 2, j - 2] = (range.Cells[i, j] as Excel.Range).Value2;
                }
            }

            workbook.Close(true, null, null);
            app.Quit();
            releaseObject(app);
        }

        private List<List<double>> LogisticPorabalisticChain(int period)
        {
            List<List<double>> result = new List<List<double>>();

            //Вспомогательный массив для дальнейшего нахождения вероятностных цепочек
            double[] sum = new double[row];

            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    sum[i] += initialData[i, j];
                }
            }

            //Массив arrPi- Pkt, т. е. вероятности
            //Массив состоит из вероятностных цепочек по 8 странам 
            //в каждый из моментов времени t
            double[,] arrPi = new double[row, col];

            for (int i = 0; i < row; i++)
                for (int j = 0; j < col; j++)
                    arrPi[i, j] = initialData[i, j] / sum[i];

            //Массив arrZ-Zkt=Pkt/P1t. Фиксируем скорость прироста 
            //по отношению к первой стране
            double[,] arrZ = new double[row, col];

            for (int i = 0; i < row; i++)
            {
                //!!!
                for (int j = 1; j < col; j++)
                {
                    arrZ[i, j] = arrPi[i, j] / arrPi[i, 0];
                }
            }

            //Массив arrMul- (Zkt+1)*(Zkt)
            double[,] arrMul = new double[row, col];

            for (int j = 1; j < col; j++)
            {
                for (int i = 0; i < row - 1; i++)
                {
                    arrMul[i, j] = arrZ[i, j] * arrZ[i + 1, j];
                }
            }

            //Массив arrMul2- Zkt^2
            double[,] arrMul2 = new double[row, col];

            for (int j = 1; j < col; j++)
            {
                for (int i = 0; i < row - 1; i++)
                {
                    arrMul2[i, j] = arrZ[i, j] * arrZ[i, j];
                }
            }

            //Массив sumMul- суммы (Zkt+1)*(Zkt) по странам
            double[] sumMul = new double[col];

            for (int j = 1; j < col; j++)
            {
                for (int i = 0; i < row; i++)
                {
                    sumMul[j] = sumMul[j] + arrMul[i, j];
                }
            }

            //Массив sumMul2- суммы Zkt^2 по странам
            double[] sumMul2 = new double[col];

            for (int j = 1; j < col; j++)
            {
                for (int i = 0; i < row; i++)
                {
                    sumMul2[j] = sumMul2[j] + arrMul2[i, j];
                }
            }

            //Массив arrY- Yk
            double[] arrY = new double[col];

            arrY[0] = 1;// пример за стандарт первую территорию/популяцию

            for (int j = 1; j < col; j++)
            {
                arrY[j] = sumMul[j] / sumMul2[j];
            }

            //Массив arrIJ- Матрица взаимного влияния
            double[,] arrIJ = new double[col, col];

            for (int i = 0; i < col; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    if (i == j)
                    {
                        arrIJ[i, j] = 0;
                    }
                    else
                    {
                        arrIJ[i, j] = 1 - arrY[j] / arrY[i];
                    }
                }
            }

            //Массив arrSumZ - сумма Zkt
            double[] arrSumZ = new double[col];

            for (int j = 0; j < col; j++)
                for (int i = 0; i < row; i++)
                {
                    arrSumZ[j] = arrSumZ[j] + arrZ[i, j];
                }

            //Массив arrMulZ - arrSumZ[j]*arrY^11
            double[] arrMulZ = new double[col];

            for (int j = 0; j < col; j++)
            {
                arrMulZ[j] = arrSumZ[j] * System.Math.Pow(arrY[j], row);
            }

            //Массив arrMultiplier 
            double[] arrMultiplier = new double[col];

            for (int j = 1; j < col; j++)
            {
                arrMultiplier[j] = (1 - arrY[j] * arrY[j]) / (1 - System.Math.Pow(arrY[j], row * 2));
            }

            //Массив arrZk0- Zk0
            double[] arrZk0 = new double[col];

            for (int j = 1; j < col; j++)
            {
                arrZk0[j] = arrMultiplier[j] * arrMulZ[j];
            }

            //P10 для нахождения начального состояния системы в терминах долей популяции
            double P10 = 0;

            for (int j = 0; j < col; j++)
            {
                P10 = P10 + arrZk0[j];
            }

            P10 = 1 / (1 + P10);

            //Начальное состояние системы в терминах долей популяции Pk0
            double[] arrPk0 = new double[col];

            for (int j = 0; j < col; j++)
            {
                arrPk0[j] = P10 * arrZk0[j];
            }

            //Интерполяция P1t
            double[] arrP1t = new double[row + period];

            for (var i = 0; i < row + period; i++)
            {
                arrP1t[i] = 0;

                for (var j = 0; j < col; j++)
                {
                    arrP1t[i] = arrP1t[i] + arrZk0[j] * System.Math.Pow(arrY[j], i);
                }

                arrP1t[i] = 1 / (1 + arrP1t[i]);

            }

            //double[,] arrInterp = new double[row + 16, col];
            List<double> tmp = new List<double>();

            for (int j = 0; j < col; j++)
            {
                for (int i = 0; i < row + period; i++)
                {
                    if (j != 0)
                        tmp.Add(arrP1t[i] * arrZk0[j] * System.Math.Pow(arrY[j], i));
                    else
                        tmp.Add(arrP1t[i]);
                }
                result.Add(tmp);
                tmp = new List<double>();
            }

            return result;
        }

        private List<List<double>> LinearLogariphmicPorabalisticChain()
        {
            List<List<double>> result = new List<List<double>>();
            List<double> tmp;

            //Вспомогательный массив для дальнейшего нахождения вероятностных цепочек
            double[] sum = new double[row];

            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    sum[i] += initialData[i, j];
                }
            }

            //Массив arrPi- Pkt, т. е. вероятности
            //Массив состоит из вероятностных цепочек по 8 странам 
            //в каждый из моментов времени t
            double[,] arrPi = new double[row, col];

            for (int i = 0; i < row; i++)
                for (int j = 0; j < col; j++)
                    arrPi[i, j] = initialData[i, j] / sum[i];

            double[,] arrHelpY = new double[row-1, col-1];
            double[,] arrY = new double[row-1, col - 1];


            int m = 0;
            int n = 1;

              for (int i = 0; i < row-1; i++)
                {

                n = 1;
                m++;
                for (int j = 0; j < col-1; j++)
                {
                    arrHelpY[i, j] =  Math.Log(arrPi[m, n])- Math.Log(arrPi[m, 0]);
                    n++;
                }
               

            }

            int p = row-1;
            int k = col-1;

            for (int i = 0; i < row - 1; i++)
            {

                k = 0;
                p--;
                for (int j = 0; j < col - 1; j++)
                {
                    arrY[i, j] = arrHelpY[p, k];
                    k++;
                }


            }

            double[,] arrHelpX = new double[row, col];
            double[,] arrX = new double[row-1, col+1];

            for (int j = 0; j < col; j++)
            {

                for (int i = 0; i < row; i++)
                {
                        arrHelpX[i, j] = Math.Log(arrPi[i, j]);
                 
                }
                
            }

            int l = 0;
            int t = row-2;

            
            for (int i = 0; i < row-1; i++)
            {
                l = 0;
                for (int j = 0; j < col+1; j++)
                {
                    

                    if (j == 0)
                    {
                        arrX[i, j] = 1;

                    }

                    else
                    {
                        arrX[i, j] = arrHelpX[t, l];
                       
                        l++;
                    }
                }
                t--;
         
            }

            Matrix<double> arrXtransp = Matrix<double>.Build.DenseOfArray(arrX);
            arrXtransp = arrXtransp.Transpose();

            Matrix<double> arrx = Matrix<double>.Build.DenseOfArray(arrX);
            Matrix<double> arry = Matrix<double>.Build.DenseOfArray(arrY);

            //!!
            Matrix<double> arrXMulti;
            arrXMulti = arrXtransp.Multiply(arrx);

            Matrix<double> arrResult = arrXMulti.Inverse() * arrXtransp * arry;


            double[,] arrA = arrResult.ToArray();
            double[,] startValue = new double[1, col-1];

            for (int j = 0; j < col-1; j++)
            {
                startValue[0, j] = Math.Exp(arrA[0, j]);
                k = 0;
                for (int i = 1; i < 9; i++)
                {
                    startValue[0, j] = startValue[0, j] * Math.Pow(arrPi[0,k], arrA[i, j]);
                    k++;
                }

            }

            double sumStart = 0;

            for (int h = 0; h < 7; h++)
            {

                sumStart = sumStart + startValue[0, h];

            }

            sumStart = 1 / (1 + sumStart);


            double[,] tmpResult = new double[row, col];
            double[,] helpResult = new double[1, col];


            for (int i = 0; i < row; i++)
            {
                if (i == 0)
                {
                    int q = 0;
                    for (int j = 0; j < col; j++)
                    {

                        
                        if (j == 0)
                        {
                            

                            tmpResult[i, 0] = sumStart;

                        }
                        else
                        {
                            tmpResult[i, j] = sumStart * startValue[i, q];
                            q++;
                        }

                    }
                }
                else
                {

                    for (int j = 0; j < col-1; j++)
                    {
                       helpResult[0, j] = Math.Exp(arrA[0, j]);
                        k = 0;

                        for (int z = 1; z < 9; z++)
                        {
                            helpResult[0, j] = helpResult[0, j] * Math.Pow(tmpResult[i-1, k], arrA[z, j]);
                            k++;
                        }

                    }

                    double helpStart = 0;

                    for (int h = 0; h < 8; h++)
                    {

                        helpStart = helpStart + helpResult[0, h];

                    }

                    helpStart = 1 / (1 + helpStart);

                        tmpResult[i, 0] = helpStart;
                       
                        for (int z = 1; z < 9; z++)
                        {
                         tmpResult[i, z] = helpStart* helpResult[0, z];
                           
                        }


                }
            
              
            }

          



            for (int i = 0; i < arrResult.RowCount; i++)
            {
                tmp = new List<double>();

                for (int j = 0; j < arrResult.ColumnCount; j++)
                {
                    tmp.Add(tmpResult[i, j]);

                }

                result.Add(tmp);

            }

            return result;
        }

        private void uploadToFile(string fileName)
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            //Значения [y - строка,x - столбец]

            for (int i = 2; i <= col + 1; i++)
            {
                ObjWorkSheet.Cells[1, i] = arrCountry[i - 2];
            }

            int tmp = 0;

            for (int j = 2; j <= result[0].Count + 1; j++)
            {
                tmp = j >= arrYears.Length ? ++tmp : arrYears[j - 2];
                ObjWorkSheet.Cells[j, 1] = tmp;
            }

            for (int j = 0; j < result.Count; j++)
            {
                for (int i = 0; i < result[j].Count; i++)
                {
                    ObjWorkSheet.Cells[i + 2, j + 2] = result[j][i];
                }
            }


            ObjWorkBook.SaveAs(fileName);


            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void logisticGrowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (row == 0 || col == 0)
            {
                MessageBox.Show("Empty excel");
                return;
            }

            int period = int.Parse(periodToolStripComboBox.SelectedText);

            result = LogisticPorabalisticChain(period);

            SolutionForm sf = new SolutionForm(arrCountry, arrYears, result);

            sf.MdiParent = this;
            sf.Show();
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                LoadFromFile(ofd.FileName);
                MessageBox.Show("Данные успешно загружены!");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void linearLogariphGrowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (row == 0 || col == 0)
            {
                MessageBox.Show("Empty excel");
                return;
            }

            result = LinearLogariphmicPorabalisticChain();

            SolutionForm sf = new SolutionForm(arrCountry, arrYears, result);

            sf.MdiParent = this;
            sf.Show();
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfg = new SaveFileDialog();
            sfg.DefaultExt = ".xls";
            sfg.Filter = "Excel|*.xls|Excel(.xlsx)|*.xlsx";

            if (result.Count>0&&result[0].Count>0)
            {
                if (sfg.ShowDialog()== DialogResult.OK)
                {
                    uploadToFile(sfg.FileName);
                }
               
            }
            
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {

        }
    }
}
