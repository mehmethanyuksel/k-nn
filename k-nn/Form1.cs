using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace k_nn
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string dosyaYolu;
        private void button1_Click(object sender, EventArgs e)
        {
            dosyaOku dosya = new dosyaOku();
            baglanti baglan = new baglanti(dosya.DosyaYolu.ToString());
            dosyaYolu = dosya.DosyaYolu;
            if (dosya.DosyaYolu != "")
            {
                baglan.xlsxbaglanti.Open();
                baglan.tablo.Clear();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Training_Data$]", baglan.xlsxbaglanti);
                da.Fill(baglan.tablo);
                dataGridView1.DataSource = baglan.tablo;
                baglan.xlsxbaglanti.Close();

                baglanti baglan2 = new baglanti(dosya.DosyaYolu.ToString());
                baglan2.xlsxbaglanti.Open();
                baglan2.tablo.Clear();
                OleDbDataAdapter da2 = new OleDbDataAdapter("SELECT * FROM [Test_Data$]", baglan2.xlsxbaglanti);
                da2.Fill(baglan2.tablo);
                dataGridView2.DataSource = baglan2.tablo;
                baglan2.xlsxbaglanti.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            int komsu = Convert.ToInt16(textBox1.Text), sayac = 0, sayac2 = 0, kayitsay = 0;
            try
            {
                baglanti baglan = new baglanti(dosyaYolu);
                baglan.xlsxbaglanti.Open();
                OleDbCommand komut = new OleDbCommand("SELECT * FROM [Training_Data$]", baglan.xlsxbaglanti);
                OleDbDataReader oku = komut.ExecuteReader();
                
                while (oku.Read())
                    sayac++;
                baglan.xlsxbaglanti.Close();
                string[,] allTrainingData = new string[sayac, 7];
                baglan.xlsxbaglanti.Open();
                OleDbDataReader oku2 = komut.ExecuteReader();
                while (oku2.Read())
                {
                    allTrainingData[kayitsay, 0] = Convert.ToString(oku2["STG"]);
                    allTrainingData[kayitsay, 1] = Convert.ToString(oku2["SCG"]);
                    allTrainingData[kayitsay, 2] = Convert.ToString(oku2["STR"]);
                    allTrainingData[kayitsay, 3] = Convert.ToString(oku2["LPR"]);
                    allTrainingData[kayitsay, 4] = Convert.ToString(oku2["PEG"]);
                    allTrainingData[kayitsay, 5] = Convert.ToString(oku2["UNS"]);
                    kayitsay++;
                }
                baglan.xlsxbaglanti.Close();


                baglanti baglan2 = new baglanti(dosyaYolu);
                baglan2.xlsxbaglanti.Open();
                OleDbCommand komut2 = new OleDbCommand("SELECT * FROM [Test_Data$]", baglan2.xlsxbaglanti);
                OleDbDataReader oku3 = komut2.ExecuteReader();
                kayitsay = 0;
                while (oku3.Read())
                    sayac2++;
                baglan2.xlsxbaglanti.Close();
                string[,] allTestData = new string[sayac2, 7];
                baglan2.xlsxbaglanti.Open();
                OleDbDataReader oku4 = komut2.ExecuteReader();
                while (oku4.Read())
                {
                    allTestData[kayitsay, 0] = Convert.ToString(oku4["STG"]);
                    allTestData[kayitsay, 1] = Convert.ToString(oku4["SCG"]);
                    allTestData[kayitsay, 2] = Convert.ToString(oku4["STR"]);
                    allTestData[kayitsay, 3] = Convert.ToString(oku4["LPR"]);
                    allTestData[kayitsay, 4] = Convert.ToString(oku4["PEG"]);
                    allTestData[kayitsay, 5] = Convert.ToString(oku4["UNS"]);
                    kayitsay++;
                }
                baglan2.xlsxbaglanti.Close();
                MessageBox.Show("Veriler başarıyla alındı.");
                hesapla hesapYap = new hesapla();
                string[,] gelenData = new string[sayac2, 7];
                int dogruSayisi;
                if(comboBox1.SelectedItem == "Öklit")
                    dogruSayisi = hesapYap.oklit(allTrainingData, allTestData, komsu, sayac,sayac2);
                else
                    dogruSayisi = hesapYap.manhattan(allTrainingData, allTestData, komsu, sayac,sayac2);

                MessageBox.Show("Test data'daki doğru veri sayısı:"+dogruSayisi+"\nYanlış veri sayısı:"+(sayac2-dogruSayisi));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu."+ex.ToString());
            }
            
            
        }

    }
}
