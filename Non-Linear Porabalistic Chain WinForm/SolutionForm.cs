using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Non_Linear_Porabalistic_Chain_WinForm
{
    public partial class SolutionForm : Form
    {
        private string[] columns;
        private int[] rows;
        private List<List<double>> Data;

        public SolutionForm(string[] columns, int[] rows, List<List<double>> Data) 
        {
            this.columns = columns;
            this.Data = Data;
            this.rows = rows;

            InitializeComponent();
        }
    }
}
