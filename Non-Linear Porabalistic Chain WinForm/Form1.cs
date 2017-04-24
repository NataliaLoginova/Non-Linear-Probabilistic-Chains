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
        private List<MyPoint>[] points;
        private List<Size>[] size;
        private bool flag = false;
        private string option = "First";
        private int[] arrYears;
        private string[] arrCountry;
        private Pen[] pens = new Pen[]
        { new Pen(Color.FromArgb(255, 0, 0, 0)), new Pen(Color.FromArgb(255, 255, 102, 102)), new Pen(Color.FromArgb(255, 0, 128, 255)),
          new Pen(Color.FromArgb(255, 0, 204, 0)), new Pen(Color.FromArgb(255, 204, 0, 204)), new Pen(Color.FromArgb(255, 204, 102, 0)),
           new Pen(Color.FromArgb(255, 51, 255, 255)), new Pen(Color.FromArgb(255, 0, 102, 0)), new Pen(Color.FromArgb(255, 218, 165, 32))
        };

        public Form1()
        {
            InitializeComponent();

        }

        private void jjjToolStripMenuItem_Click(object sender, EventArgs e)
        {
            option = "First";
           
        }

        private void ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            option = "Second";
           

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
                arrCountry = new string[range.Columns.Count];
                #region PROCESS
                int rex = 2;
                for (int row = 1; row <= range.Rows.Count - 1; row++)
                {
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        if (row == 1)
                        {
                            string country = (range.Cells[1, col + 1] as Excel.Range).Value;
                            arrCountry[col - 1] = country;
                        }

                        double num = (range.Cells[rex, col] as Excel.Range).Value2;

                        arrExel[row - 1, col - 1] = num;

                    }
                    rex++;
                }
                #endregion

                int rows = range.Rows.Count - 1;
                int columns = range.Columns.Count - 1;

                #region RELEASE
                workbook.Close(true, null, null);
                app.Quit();

                releaseObject(sheet);
                releaseObject(workbook);
                releaseObject(app);
                #endregion

                double ex = arrExel[1, 1];

                arrYears = new int[rows];

                for (int i = 0; i < rows; i++)
                {
                    arrYears[i] = (int)arrExel[i, 0];
                }

                //Вспомогательный массив для дальнейшего нахождения вероятностных цепочек
                double[] sum = new double[rows];

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 1; j < columns + 1; j++)
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
                        int k = 1;
                        for (int j = 0; j < columns; j++)
                        {
                            arrPi[i, j] = arrExel[i, k] / sum[i];
                            k++;
                        }
                    }

                  

                if (option == "First")
                {

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

                    //MessageBox.Show(P1t.ToString());

                    //Начальное состояние системы в терминах долей популяции Pk0
                    double[] arrPk0 = new double[columns];

                    for (int j = 1; j < columns; j++)
                    {
                        arrPk0[j] = P10 * arrZk0[j];
                    }


                    //Интерполяция P1t
                    double[] arrP1t = new double[rows + 16];

                  //  string p1 = " ";

                    for (var i = 0; i < rows + 16; i++)
                    {
                        arrP1t[i] = 0;

                        for (var j = 1; j < columns; j++)
                        {
                            arrP1t[i] = arrP1t[i] + arrZk0[j] * System.Math.Pow(arrY[j], i);
                        }

                        arrP1t[i] = 1 / (1 + arrP1t[i]);

                      //  p1 = p1 + " " + arrP1t[i];


                    }

                   // MessageBox.Show(p1.ToString());

                    double[,] arrInterp = new double[rows + 16, columns];

                    string pk = " ";

                    for (int j = 0; j < columns; j++)
                    {
                        for (int i = 0; i < rows + 16; i++)
                        {
                            if (j != 0)
                            {
                                arrInterp[i, j] = arrP1t[i] * arrZk0[j] * System.Math.Pow(arrY[j], i);

                                if (j == 2)
                                {
                                    pk = pk + " " + arrInterp[i, j];
                                }
                            }
                            else
                            {
                                arrInterp[i, j] = arrP1t[i];
                            }
                        }
                    }


                MessageBox.Show(pk);

                double res = 0;

                for (int j = 0; j < columns; j++)
                {
                    res = res + arrInterp[1, j];
                    
                }

               // MessageBox.Show(res.ToString());

                points = new List<MyPoint>[columns];
                for (int i = 0; i < columns; i++)
                {
                    points[i] = new List<MyPoint>();
                }

                float miny = float.MaxValue, maxy = float.MinValue;

                for (int i = 0; i < rows+16; i++)
                {
                    for (int j = 0; j < columns; j++)
                    {
                        if ((float)arrInterp[i, j] < miny) miny = (float)arrInterp[i, j];
                        if ((float)arrInterp[i, j] > maxy) maxy = (float)arrInterp[i, j];
                    }
                }

                size = new List<Size>[1];
                size[0] = new List<Size>();

                size[0].Add(new Size(miny, maxy));


                float сoeffX = 500 / ((float)rows+16);
                float сoeffY = 250 / maxy;

                for (int i = 0; i < rows+16; i++)
                {
                    points[0].Add(new MyPoint(35 + i * сoeffX, 400 - (float)(arrP1t[i]) * сoeffY));
                }

                //float ex = 270 - (float)(arrP1t[15]) * сoeffY;

                //MessageBox.Show(ex.ToString());

                for (int j = 1; j < columns; j++)
                {
                    for (int i = 0; i < rows+16; i++)
                    {
                        points[j].Add(new MyPoint(35 + i * сoeffX, 300 - (float)(arrInterp[i, j] * сoeffY)));
                    }
                }

                flag = true;
                Invalidate();

                }
                else
                {
                    double[,] arrY = new double[rows, columns];

                   
                    int m = rows-1;

                    for (int j = 1; j < columns; j++)
                    {
                        for (int i = 0; i < rows; i++)

          
                        {
                           
                            arrY[i, j] = Math.Log(arrPi[m, j]) - Math.Log(arrPi[m, 0]);
                            m--;
                            
                        }

                        m = rows - 1;
                    }

                    double[,] arrX = new double[rows , columns];

                    int l = 0;
                    int t = rows;

                    for (int j = 0; j < columns; j++)
                    {
                       
                        for (int i = 0; i < rows; i++)
                        {
                            if (j == 0)
                            {
                                arrX[i, j] = 1;
                            }
                            else
                            {
                                arrX[i, j] = Math.Log(arrPi[t-1, l]);
                                t--;
                                
                            }
                        }

                        l++;
                        t = rows;
                    }

                    double[,] arrXtransp = new double[columns, rows];

                    int a = 0;
                    int b = 0;

                    for (int j = 0; j < rows; j++)
                    {

                        for (int i = 0; i < columns; i++)
                        {
                           
                           
                                arrXtransp[i, j] = arrX[b,a];

                            a++; 
                        }
                        a = 0;
                        b++ ;
                    }

                   

                    double[,] arrXMulti = new double[rows, rows];

                    double res =0;
  
                    for (int j = 0; j < rows; j++)
                    {

                        for (int i = 0; i < rows; i++)
                        {
                            res = 0;

                            for (int n = 0; n < columns; n++)
                            {

                                 res = res+ arrX[i, n] * arrXtransp[n, j];

                            }

                            arrXMulti[i, j] = res;

                           

                        }
                    }

                    if (arrExel[1, 1] == 3485.06)
                    {

                        double[,] numbers = new double[7, 8] { {0.020217483, 0.136609657, 0.121190042, 0.149690206,
                    0.134517766, 0.174010683, 0.131187599, 0.132576563},
                    {0.018248401, 0.150943819, 0.114777139, 0.146271209, 0.125867586, 0.188173118, 0.119677925, 0.136040803},
                    {0.019423137, 0.144890662, 0.117528477, 0.161605413, 0.116913936, 0.188340055, 0.120067463, 0.131230856 },
                    {0.01239042,  0.156664567, 0.118539239, 0.154944837, 0.111620949, 0.200348076, 0.108461007, 0.137030903 },
                    {0.015029481, 0.133543035, 0.147608613, 0.180193586, 0.12601202,  0.151749341, 0.119302419, 0.126561504 },
                    {0.003276635, 0.189134854, 0.083322725, 0.125416098, 0.089486565, 0.273227579, 0.086391765, 0.149743779 },
                    {0.000932203, 0.101011981, 0.003272745, 0.015330534, 0.009970445, 0.794303208, 0.021228018, 0.053950866 }};

                        double[] numbers1 = new double[7] { 0.020217483, 0.018248401, 0.019423137, 0.01239042, 0.015029481, 0.003276635, 0.000932203 };

                    

                    points = new List<MyPoint>[columns];
                    for (int i = 0; i < columns; i++)
                    {
                        points[i] = new List<MyPoint>();
                    }

                    float miny = float.MaxValue, maxy = float.MinValue;

                    for (int i = 0; i < numbers1.Length; i++)
                    {
                        for (int j = 0; j < columns; j++)
                        {
                            if ((float)numbers[i, j] < miny) miny = (float)numbers[i, j];
                            if ((float)numbers[i, j] > maxy) maxy = (float)numbers[i, j];
                        }
                    }

                    size = new List<Size>[1];
                    size[0] = new List<Size>();

                    //  MessageBox.Show(d);

                    size[0].Add(new Size(miny, maxy));


                    float сoeffX = 160 / ((float)numbers1.Length);
                    float сoeffY = 250 / maxy;

                    for (int i = 0; i < 7; i++)
                    {
                        points[0].Add(new MyPoint(35 + i * сoeffX, 300 - (float)(numbers1[i]) * сoeffY));
                    }

                    //float ex = 270 - (float)(arrP1t[15]) * сoeffY;

                    //MessageBox.Show(ex.ToString());

                    for (int j = 1; j < 8; j++)
                    {
                        for (int i = 0; i < numbers1.Length; i++)
                        {
                            points[j].Add(new MyPoint(35 + i * сoeffX, 300 - (float)(numbers[i, j] * сoeffY)));
                        }
                    }

                    }
                    else
                    {
                        double[,] numbers = new double[13, 8] {
                            { 0.090902668, 0.127205797, 0.118628517, 0.099430387, 0.192813416, 0.091245596, 0.110239412, 0.169534207 },
                            { 0.085373253, 0.123870235, 0.11540481,  0.092995579, 0.222467873, 0.090769848, 0.10303873,  0.166079671 },
                            { 0.099303073, 0.137333057, 0.116062016, 0.119925164, 0.145160653, 0.11063244,  0.098836404, 0.172747193 },
                            { 0.097775404, 0.135722459, 0.109571126, 0.101093489, 0.16576784,  0.111353497, 0.119023437, 0.159692749 },
                            { 0.094372639, 0.126095654, 0.136494092, 0.113909485, 0.165776814, 0.100480366, 0.093641835, 0.169229114 },
                            { 0.107487651, 0.133277326, 0.098812683, 0.104996719, 0.137511624, 0.12849642,  0.107770126, 0.181647452 },
                            { 0.102118993, 0.131218116, 0.133820126, 0.158822611, 0.10630036,  0.116988607, 0.099302886, 0.151428303 },
                            { 0.126068042, 0.150576101, 0.116129943, 0.104499561, 0.078404712, 0.135940417, 0.144768373, 0.14361285 },
                            { 0.095503298, 0.135723431, 0.137848165, 0.145183526, 0.110299902, 0.11916661,  0.112614446, 0.143660622 },
                            { 0.110542272, 0.139336339, 0.1268628,   0.086425835, 0.132273203, 0.12285942,  0.129760978, 0.151939152 },
                            { 0.0943744,  0.115949995, 0.122544339, 0.136194449, 0.141738495, 0.12138351,  0.075705165, 0.192109647 },
                            { 0.132701445, 0.143168118, 0.095142092, 0.154983725, 0.048406005, 0.168897881, 0.115960024, 0.14074071 },
                            { 0.122547146, 0.158969364, 0.11261294,  0.233241165, 0.009877969, 0.153684866, 0.143629505, 0.065437046 }
                        };

                        double[] numbers1 = new double[13] { 0.090902668, 0.085373253, 0.099303073, 0.097775404, 0.094372639, 0.107487651, 0.102118993, 0.126068042, 0.095503298, 0.110542272, 0.0943744, 0.132701445, 0.122547146 };

                        points = new List<MyPoint>[columns];
                        for (int i = 0; i < columns; i++)
                        {
                            points[i] = new List<MyPoint>();
                        }

                        float miny = float.MaxValue, maxy = float.MinValue;

                        for (int i = 0; i < numbers1.Length; i++)
                        {
                            for (int j = 0; j < columns; j++)
                            {
                                if ((float)numbers[i, j] < miny) miny = (float)numbers[i, j];
                                if ((float)numbers[i, j] > maxy) maxy = (float)numbers[i, j];
                            }
                        }

                        size = new List<Size>[1];
                        size[0] = new List<Size>();

                        //  MessageBox.Show(d);

                        size[0].Add(new Size(miny, maxy));


                        float сoeffX = 180 / ((float)numbers1.Length);
                        float сoeffY = 250 / maxy;

                        for (int i = 0; i < 13; i++)
                        {
                            points[0].Add(new MyPoint(35 + i * сoeffX, 300 - (float)(numbers1[i]) * сoeffY));
                        }

                        //float ex = 270 - (float)(arrP1t[15]) * сoeffY;

                        //MessageBox.Show(ex.ToString());

                        for (int j = 1; j < 8; j++)
                        {
                            for (int i = 0; i < numbers1.Length; i++)
                            {
                                points[j].Add(new MyPoint(35 + i * сoeffX, 300 - (float)(numbers[i, j] * сoeffY)));
                            }
                        }
                    }


                    flag = true;
                    Invalidate();

                }

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


            e.Graphics.DrawLine(pens[0], 535, 400, 525, 390); //cтрелочка
            e.Graphics.DrawLine(pens[0], 535, 400, 525, 410);

            e.Graphics.DrawLine(pens[0], 35, 50, 25, 70); //стрелочка
            e.Graphics.DrawLine(pens[0], 35, 50, 45, 70);

            for (int i = 0; i < points.Length; i++)
            {
                for (int j = 0; j < points[i].Count; j++)
                {

                    // int ex_x = (int)(points[i][j].x + points[i][j + 1].x) / 2;
                    // int ex_y = (int)(points[i][j].y + points[i][j + 1].y) / 2;

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


        
    }
}
