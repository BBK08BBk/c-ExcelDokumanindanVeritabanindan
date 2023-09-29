using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Data.SqlClient;
using static System.Runtime.InteropServices.JavaScript.JSType;
using excel = Microsoft.Office.Interop.Excel;

namespace ExcelVtEntegrasyonProje
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        SqlConnection conn = new SqlConnection(@"Data Source=BBK\SQLEXPRESS;Initial Catalog=BBK;Integrated Security=True");
        //Data Source = BBK\SQLEXPRESS;Initial Catalog = BBK; Integrated Security = True
        private void btnvtdenoku_Click(object sender, EventArgs e)
        {
            excel.Application application = new excel.Application();
            application.Visible = true;
            excel.Workbook wb = application.Workbooks.Add(System.Reflection.Missing.Value);
            excel.Worksheet ws = wb.Sheets[1];
            string[] basliklar = { "Personel No", "Ad", "Soyad", "Semt", "Şehir" };
            Microsoft.Office.Interop.Excel.Range range;
            for (int i = 0; i < basliklar.Length; i++)
            {
                range = ws.Cells[1, 1 + i];
                range.Value2 = basliklar[i];
            }


            try
            {

                conn.Open();
                string sqlCumle = "Select * From Personel";
                SqlCommand cmd = new SqlCommand(sqlCumle, conn);
                SqlDataReader dr = cmd.ExecuteReader();
                int satir = 2;
                while (dr.Read())
                {
                    string PersNo = dr[0].ToString();
                    string ad = dr[1].ToString();
                    string soyad = dr[2].ToString();
                    string semt = dr[3].ToString();
                    string sehır = dr[4].ToString();
                    richTextBox1.Text = richTextBox1.Text + PersNo + "-" + "-" +
                        ad + "-" + "-" + soyad + "-" + "-" + semt + "-" + "-" + sehır + "\n";

                    range = ws.Cells[satir, 1];
                    range.Value2 = PersNo;
                    range = ws.Cells[satir, 2];
                    range.Value2 = ad;
                    range = ws.Cells[satir, 3];
                    range.Value2 = soyad;
                    range = ws.Cells[satir, 4];
                    range.Value2 = semt;
                    range = ws.Cells[satir, 5];
                    range.Value2 = sehır;
                    satir++;
                }

            }
            catch (Exception b)
            {
                Console.WriteLine(b.Message.ToString());
            }
            finally
            {
                if (conn != null)
                {
                    conn.Close();
                }
            }

        }

        private void btnexceloku_Click(object sender, EventArgs e)
        {
            excel.Application eapp;
            excel.Workbook ewb;
            excel.Worksheet ews;
            excel.Range erange;
            int Rcount = 0;
            int cCount = 0;
            eapp = new excel.Application();
            ewb = eapp.Workbooks.Open("C:\\Users\\berkk\\OneDrive\\Masaüstü\\txtexcel.xlsx");
            ews = ewb.Worksheets.get_Item(1);//sheet i açtık
            erange = ews.UsedRange;
            //ilk olarak richtextbox içini temizleyelim
            richTextBox2.Clear();
            //ilk satır başlık içerdiği için Rcount 2 den başlatıcaz
            for (Rcount = 2; Rcount <= erange.Rows.Count; Rcount++)
            {
                ArrayList arrayList = new ArrayList();
                for (cCount = 1; cCount <= erange.Columns.Count; cCount++)
                {
                    string okunanHucre = Convert.ToString((erange.Cells[Rcount, cCount] as excel.Range).Value2);
                    richTextBox2.Text = richTextBox2.Text + okunanHucre + " " + " ";
                    arrayList.Add(okunanHucre);
                }
                richTextBox2.Text = richTextBox2.Text + "\n";

                try
                {
                    conn.Open();
                    SqlCommand sqlCommand = new SqlCommand("Insert into Personel(PersonelNo,Ad,Soyad,Semt,Sehır) "
                        + "Values (@P1,@P2,@P3,@P4,@P5)", conn);
                    sqlCommand.Parameters.AddWithValue("@P1", arrayList[0]);
                    sqlCommand.Parameters.AddWithValue("@P2", arrayList[1]);
                    sqlCommand.Parameters.AddWithValue("@P3", arrayList[2]);
                    sqlCommand.Parameters.AddWithValue("@P4", arrayList[3]);
                    sqlCommand.Parameters.AddWithValue("@P5", arrayList[4]);
                    sqlCommand.ExecuteNonQuery();

                }
                catch (Exception exx)
                {
                    Console.WriteLine(exx.Message.ToString());
                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                }

            }
            eapp.Quit();
            ReleaseObject(ews);
            ReleaseObject(ewb);
            ReleaseObject(eapp);

        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}