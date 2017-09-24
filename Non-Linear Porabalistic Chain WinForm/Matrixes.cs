using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Non_Linear_Porabalistic_Chain_WinForm
{
    class Matrixes
    {
       
        public string name;
        
        public Matrixes()
        {
            name = "New matrics";
        }

        public Matrixes(string nm)
        {
            name = nm;
        }

        public double[,] Transpose(double[,] arr, int row, int col)
        {
            double[,] arrTranspose = new double[col, row];

            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
            
                    arrTranspose[j, i] = arr[i, j]; ;
                }
            }

            return arrTranspose;
        }

        public double[,] Multi(double[,] arr, int row, int col)
        {
            double[,] arrMulti = new double[col, row];

            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {

                    arrTranspose[j, i] = arr[i, j]; ;
                }
            }

            return arrTranspose;
        }
    }
}
