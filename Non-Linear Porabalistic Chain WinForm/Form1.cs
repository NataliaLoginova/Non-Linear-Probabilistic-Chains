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



namespace Non_Linear_Porabalistic_Chain_WinForm
{


    struct MyPoint
    {
        public float x;
        public float y;

        public MyPoint(float x, float y)
        {
            this.x = x;
            this.y = y;
        }
    }
    struct Size
    {
        public float x;
        public float y;

        public Size(float x, float y)
        {
            this.x = x;
            this.y = y;
        }
    }
    public partial class Form1 : Form
    {
        private double[,] initialData;
        private int[] arrYears;
        private string[] arrCountry;
        private int col;
        private int row;
        private List<MyPoint>[] points;
        private List<Size>[] size;
        private bool flag = false;
        private string option = "First";
        private Pen[] pens = new Pen[]
        { new Pen(Color.FromArgb(255, 0, 0, 0)), new Pen(Color.FromArgb(255, 255, 102, 102)), new Pen(Color.FromArgb(255, 0, 128, 255)),
          new Pen(Color.FromArgb(255, 0, 204, 0)), new Pen(Color.FromArgb(255, 204, 0, 204)), new Pen(Color.FromArgb(255, 204, 102, 0)),
           new Pen(Color.FromArgb(255, 51, 255, 255)), new Pen(Color.FromArgb(255, 0, 102, 0)), new Pen(Color.FromArgb(255, 218, 165, 32))
        };


        public Form1()
        {
            InitializeComponent();
            col = 0;
            row = 0;

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
        }


        private void jjjToolStripMenuItem_Click(object sender, EventArgs e)
        {
            option = "First";

        }

        private void ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            option = "Second";


        }


        private List<List<double>> LogisticPorabalisticChain()
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
            double[] arrP1t = new double[row + 16];

            for (var i = 0; i < row + 16; i++)
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
                for (int i = 0; i < row + 16; i++)
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

            double[,] arrY = new double[row, col-1];


              for (int i = 0; i < row; i++)
                {
                for (int j = 0; j < col - 1; j++)
                {
                    arrY[i, j] =  Math.Log(arrPi[i, j+1])- Math.Log(arrPi[i, 0]);

                }

            }

            double[,] arrX = new double[row, col];

            int l = 0;
            int t = row;

            for (int j = 0; j < col; j++)
            {

                for (int i = 0; i < row; i++)
                {
                    if (j == 0)
                    {
                        arrX[i, j] = 1;
                    }
                    else
                    {
                        arrX[i, j] = Math.Log(arrPi[t - 1, l]);
                        t--;

                    }
                }

                l++;
                t = row;
            }

            Matrix<double> arrXtransp = Matrix<double>.Build.DenseOfArray(arrX);
            arrXtransp = arrXtransp.Transpose();

            Matrix<double> arrx = Matrix<double>.Build.DenseOfArray(arrX);
            Matrix<double> arry = Matrix<double>.Build.DenseOfArray(arrY);

            //!!
            Matrix<double> arrXMulti;
            arrXMulti = arrXtransp.Multiply(arrx);

            Matrix<double> arrResult = arrXMulti.Inverse() * arrXtransp * arry;


            double[,] tmpResult = arrResult.ToArray();
            

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




        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            if (!flag) return;

            Image newImageLine = Image.FromFile("lines.png");
            Image newImageCricle = Image.FromFile("cricle.png");
            Image newImagePoints = Image.FromFile("points.png");
            Image newImageLinePoint = Image.FromFile("line-point.png");
            Image newImageCross = Image.FromFile("cross.png");
            Image newImageTriangle = Image.FromFile("triangle.png");
            Image newImageRhombus = Image.FromFile("rhombus.png");
            Image newImagePointLines = Image.FromFile("point-line.png");

            Image[] arrImages = new Image[8];
            arrImages[0] = newImageLine;
            arrImages[1] = newImageCricle;
            arrImages[2] = newImagePoints;
            arrImages[3] = newImageLinePoint;
            arrImages[4] = newImageCross;
            arrImages[5] = newImageTriangle;
            arrImages[6] = newImageRhombus;
            arrImages[7] = newImagePointLines;


            e.Graphics.DrawLine(pens[0], 35, 400, 535, 400);  //ось Ox
            string nameOx = "Time (Year)";
            e.Graphics.DrawString(nameOx.ToString(),
           new Font("Arial", 10), System.Drawing.Brushes.Black, new Point(525, 410));

            e.Graphics.DrawLine(pens[0], 35, 50, 35, 400); //oсь Oy
            string nameOy = "P (probability)";
            e.Graphics.DrawString(nameOy.ToString(),
           new Font("Arial", 10), System.Drawing.Brushes.Black, new Point(10, 48));


            e.Graphics.DrawLine(pens[0], 535, 400, 525, 390);//cтрелочка
            e.Graphics.DrawLine(pens[0], 535, 400, 525, 410);

            e.Graphics.DrawLine(pens[0], 35, 50, 25, 70); //стрелочка
            e.Graphics.DrawLine(pens[0], 35, 50, 45, 70);

            for (int i = 0; i < points.Length; i++)
            {
                for (int j = 0; j < points[i].Count; j++)
                {

                    e.Graphics.DrawImage(arrImages[i], new Point((int)points[i][j].x, (int)points[i][j].y));

                }
            }

            float сoeffY = size[0][0].y / 12;

            for (int i = 1; i < 13; i++)
            {
                double value = Math.Round(сoeffY * i, 3);
                if (i == 11)
                {
                    e.Graphics.DrawString(value.ToString(),
                new Font("Arial", 10), System.Drawing.Brushes.Blue, new Point(0, 390 - 30 * i));

                }
                else
                {
                    e.Graphics.DrawLine(pens[0], 30, 400 - 30 * i, 40, 400 - 30 * i); //подписи для оси Oy
                    e.Graphics.DrawString(value.ToString(),
                    new Font("Arial", 10), System.Drawing.Brushes.Blue, new Point(0, 390 - 30 * i));
                }
            }



            double count = Math.Round(((double)arrYears.Length + 16) / 10);

            float step = (float)(500 * count) / ((float)arrYears.Length + 16);
            int index = 0;
            int year = arrYears[arrYears.Length - 1];

            //.Show(step.ToString());

            for (int i = 1; i < 11; i++)
            {
                index = index + (int)count;

                if (index < arrYears.Length + 16)
                {
                    if (index < arrYears.Length)
                    {
                        e.Graphics.DrawLine(pens[0], 35 + (int)step * i, 395, 35 + (int)step * i, 405); //подписи для оси Ох
                        e.Graphics.DrawString(arrYears[index].ToString(),
            new Font("Arial", 10), System.Drawing.Brushes.Blue, new Point((int)(20 + (int)step * i), 405));
                    }
                    else
                    {
                        year = year + (int)count;
                        e.Graphics.DrawLine(pens[0], 35 + (int)step * i, 395, 35 + (int)step * i, 405);
                        e.Graphics.DrawString(year.ToString(),
            new Font("Arial", 10), System.Drawing.Brushes.Blue, new Point((int)(20 + (int)step * i), 405));
                    }
                }

            }

            for (int i = 0; i < 8; i++)
            {
                int penI = i % pens.Length + 1;
                e.Graphics.DrawImage(arrImages[i], new Point(560, 20 + 25 * i)); //легенда
                e.Graphics.DrawString(arrCountry[i].ToString(),
       new Font("Arial", 8), System.Drawing.Brushes.Black, new Point(600, 15 + 25 * i));
            }
        }

        private void logisticGrowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (row == 0 || col == 0)
            {
                MessageBox.Show("Empty excel");
                return;
            }

            

            SolutionForm sf = new SolutionForm(arrCountry, arrYears, LogisticPorabalisticChain());

            sf.MdiParent = this;
            sf.Show();
        }

        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                LoadFromFile(ofd.FileName);
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

            SolutionForm sf = new SolutionForm(arrCountry, arrYears, LinearLogariphmicPorabalisticChain());

            sf.MdiParent = this;
            sf.Show();
        }
    }
}
