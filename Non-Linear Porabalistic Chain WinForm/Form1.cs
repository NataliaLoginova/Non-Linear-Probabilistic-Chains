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
    public partial class Form1 : Form
    {
        private List<MyPoint>[] points;
        private bool flag = false;
        private Pen[] pens = new Pen[]
        { new Pen(Color.FromArgb(255, 0, 0, 0)), new Pen(Color.FromArgb(255, 255, 102, 102)), new Pen(Color.FromArgb(255, 0, 128, 255)),
            new Pen(Color.FromArgb(255, 0, 204, 0)), new Pen(Color.FromArgb(255, 204, 0, 204)),
            new Pen(Color.FromArgb(255, 204, 102, 0)), new Pen(Color.FromArgb(255, 51, 255, 255))
        };

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                #region OPEN
                Excel.Application app = new Excel.Application();
                Excel.Workbook workbook = app.Workbooks.Open(ofd.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
                Excel.Range range = sheet.UsedRange;
                #endregion

                double[,] arrExel = new double[range.Rows.Count, range.Columns.Count];
                #region PROCESS
                for (int row = 1; row <= range.Rows.Count; row++)
                {
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        double num = (range.Cells[row, col] as Excel.Range).Value2;

                        arrExel[row - 1, col - 1] = num;
                    }
                }
                #endregion

                //  Console.WriteLine(arrExel[0, 0]);
                int rows = range.Rows.Count;
                int columns = range.Columns.Count;

                #region RELEASE
                workbook.Close(true, null, null);
                app.Quit();

                releaseObject(sheet);
                releaseObject(workbook);
                releaseObject(app);
                #endregion

                //Вспомогательный массив для дальнейшего нахождения вероятностных цепочек
                double[] sum = new double[rows];

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < columns; j++)
                    {
                        sum[i] += arrExel[i, j];
                    };

                };

                //Массив arrPi- Pkt, т. е. вероятности
                //Массив состоит из вероятностных цепочек по 8 странам 
                //в каждый из моментов времени t
                double[,] arrPi = new double[rows, columns];

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < columns; j++)
                    {
                        arrPi[i, j] = arrExel[i, j] / sum[i];
                    }
                }

                //Массив arrZ-Zkt=Pkt/P1t. Фиксируем скорость прироста 
                //по отношению к первой стране
                double[,] arrZ = new double[rows, columns];

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 1; j < columns; j++)
                    {
                        arrZ[i, j] = arrPi[i, j] / arrPi[i, 0];
                    }
                }

                //Массив arrMul- (Zkt+1)*(Zkt)
                double[,] arrMul = new double[rows, columns];

                for (int j = 1; j < columns; j++)
                {
                    for (int i = 0; i < rows - 1; i++)
                    {
                        arrMul[i, j] = arrZ[i, j] * arrZ[i + 1, j];
                    }
                }

                //Массив arrMul2- Zkt^2
                double[,] arrMul2 = new double[rows, columns];

                for (int j = 1; j < columns; j++)
                {
                    for (int i = 0; i < rows - 1; i++)
                    {
                        arrMul2[i, j] = arrZ[i, j] * arrZ[i, j];
                    }
                }

                //Массив sumMul- суммы (Zkt+1)*(Zkt) по странам
                double[] sumMul = new double[columns];

                for (int j = 1; j < columns; j++)
                {
                    for (int i = 0; i < rows; i++)
                    {
                        sumMul[j] = sumMul[j] + arrMul[i, j];
                    }
                }

                //Массив sumMul2- суммы Zkt^2 по странам
                double[] sumMul2 = new double[columns];

                for (int j = 1; j < columns; j++)
                {
                    for (int i = 0; i < rows; i++)
                    {
                        sumMul2[j] = sumMul2[j] + arrMul2[i, j];
                    }
                }

                //Массив arrY- Yk
                double[] arrY = new double[columns];

                arrY[0] = 1;// пример за стандарт первую территорию/популяцию

                for (int j = 1; j < columns; j++)
                {
                    arrY[j] = sumMul[j] / sumMul2[j];
                }

                //Массив arrIJ- Матрица взаимного влияния
                double[,] arrIJ = new double[columns, columns];

                for (int i = 0; i < columns; i++)
                {
                    for (int j = 0; j < columns; j++)
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
                double[] arrSumZ = new double[columns];

                for (int j = 1; j < columns; j++)
                    for (int i = 0; i < rows; i++)
                    {
                        {
                            arrSumZ[j] = arrSumZ[j] + arrZ[i, j];
                        }
                    }

                //Массив arrMulZ - arrSumZ[j]*arrY^11
                double[] arrMulZ = new double[columns];

                for (int j = 1; j < columns; j++)
                {
                    arrMulZ[j] = arrSumZ[j] * System.Math.Pow(arrY[j], rows);
                }

                //Массив arrMultiplier 
                double[] arrMultiplier = new double[columns];

                for (int j = 1; j < columns; j++)
                {
                    arrMultiplier[j] = (1 - arrY[j] * arrY[j]) / (1 - System.Math.Pow(arrY[j], rows * 2));
                }

                //Массив arrZk0- Zk0
                double[] arrZk0 = new double[columns];

                for (int j = 1; j < columns; j++)
                {
                    arrZk0[j] = arrMultiplier[j] * arrMulZ[j];
                }

                //P10 для нахождения начального состояния системы в терминах долей популяции
                double P10 = 0;

                for (int j = 1; j < columns; j++)
                {
                    P10 = P10 + arrZk0[j];
                }

                P10 = 1 / (1 + P10);

                //P1t для интерполяции
                double P1t = 0;

                for (int j = 1; j < columns; j++)
                {
                    P1t = P1t + arrZk0[j] * System.Math.Pow(arrY[j], rows);
                }

                P1t = 1 / (1 + P1t);

                //Console.WriteLine(P1t);
                //MessageBox.Show(P1t.ToString());

                //Начальное состояние системы в терминах долей популяции Pk0
                double[] arrPk0 = new double[columns];

                for (int j = 1; j < columns; j++)
                {
                    arrPk0[j] = P10 * arrZk0[j];
                }


                //Интерполяция Pkt
                double[] arrPkt = new double[columns];

                for (int j = 1; j < columns; j++)
                {
                    arrPkt[j] = P1t * arrZk0[j] * System.Math.Pow(arrY[j], rows);
                }

                double[] arrP1 = new double[rows];

                for (var i = 0; i < rows; i++)
                {
                    arrP1[i] = 0;

                    for (var j = 1; j < columns; j++)
                    {
                        arrP1[i] = arrP1[i] + arrZk0[j] * System.Math.Pow(arrY[j], i);
                    }

                    arrP1[i] = 1 / (1 + arrP1[i]);

                }

                double[,] arrInterp = new double[rows, columns];

                for (int j = 0; j < columns; j++)
                {
                    for (int i = 0; i < rows; i++)
                    {
                        if (j != 0)
                        {
                            arrInterp[i, j] = P1t * arrZk0[j] * System.Math.Pow(arrY[j], i);
                        }
                        else
                        {
                            arrInterp[i, j] = arrP1[i];
                        }
                    }
                }

                points = new List<MyPoint>[columns];
                for (int i = 0; i < columns; i++)
                {
                    points[i] = new List<MyPoint>();
                }

                for (int j = 1; j < columns; j++)
                {
                    for (int i = 0; i < rows; i++)
                    {
                        points[j].Add(new MyPoint(50 + i * 6, 250 - (float)(arrInterp[i, j]) * 1000));
                    }
                }

                // Console.WriteLine(arrPkt[1]);
                flag = true;
                Invalidate();
            }
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
            if (flag)
            {

                for (int i = 0; i < points.Length; i++)
                {
                    for (int j = 0; j < points[i].Count - 1; j++)
                    {
                        int penI = i % pens.Length;
                        e.Graphics.DrawLine(pens[penI], points[i][j].x, points[i][j].y,
                            points[i][j + 1].x, points[i][j + 1].y);
                    }
                }

                float minx = float.MaxValue, maxx = float.MinValue;
                float miny = float.MaxValue, maxy = float.MinValue;
                for (int i = 0; i < points.Length; i++)
                {
                    for (int j = 0; j < points[i].Count; j++)
                    {

                        if (points[i][j].x < minx) minx = points[i][j].x;
                        if (points[i][j].x > maxx) maxx = points[i][j].x;
                        if (points[i][j].y < miny) miny = points[i][j].y;
                        if (points[i][j].y > maxy) maxy = points[i][j].y;
                    }
                }

                e.Graphics.DrawLine(pens[0], minx, maxy, maxx, maxy);
                e.Graphics.DrawLine(pens[0], minx, miny, minx, maxy);
            }
        }
    }
}
