using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace k_nn
{
    class baglanti
    {
        public OleDbConnection xlsxbaglanti;
        public DataTable tablo;
        public baglanti(string dosyaYolu)
        {
            xlsxbaglanti = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dosyaYolu + "; Extended Properties='Excel 12.0 Xml;HDR=YES'");
            tablo = new DataTable();
        }
    }
}
