using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Non_Linear_Porabalistic_Chain_WinForm
{
    public partial class SolutionForm : Form
    {
        private string[] columns;
        private int[] rows;
        private List<List<double>> Data;
        private String[] arrImages;
    

        public SolutionForm(string[] columns, int[] rows, List<List<double>> Data) 
        {
            this.columns = columns;
            this.Data = Data;
            this.rows = rows;

            arrImages = new String[8];
            arrImages[0] = "lines.png";
            arrImages[1] = "cricle.png";
            arrImages[2] = "points.png";
            arrImages[3] = "line-point.png";
            arrImages[4] = "cross.png";
            arrImages[5] = "triangle.png";
            arrImages[6] = "rhombus.png";
            arrImages[7] = "point-line.png";

            InitializeComponent();
        }

        private void SolutionForm_Load(object sender, EventArgs e)
        {
            int[] r = new int[Data[0].Count];
            int tmp = 0;

            for (int j = 0; j < Data[0].Count; j++)
            {
                tmp = j >= rows.Length ? ++tmp : rows[j];
                r[j] = tmp;
            }

            for (int i = 0; i<columns.Length; i++)
            {
                chart1.Series.Add(columns[i]);
                chart1.Series[i].ChartType = SeriesChartType.Line;
                chart1.Series[i].MarkerImage = arrImages[i % 8];
                chart1.Series[i].Color = Color.Black;

                for (int j = 0; j < Data[i].Count; j++)
                {
                    chart1.Series[i].Points.AddXY(r[j], Data[i][j]);
                }

            }
       }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void chart1_AxisViewChanged(object sender, ViewEventArgs e)
        {
            chart1.ChartAreas[0].RecalculateAxesScale();
        }
        
    }
}
