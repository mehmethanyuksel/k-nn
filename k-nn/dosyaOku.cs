using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace k_nn
{
    class dosyaOku
    {
        public string DosyaYolu;
        public dosyaOku()
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.ShowDialog();
            DosyaYolu = dosya.FileName;
        }
    }
}
