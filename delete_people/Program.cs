using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using SaUbiNet = Com.Saperion.UBI;

namespace delete_people
{
    class Program
    {
        public static SqlConnection SQLConn = new SqlConnection();
        public static SaUbiNet.Application oSAPApp = new SaUbiNet.Application();
        public static SaUbiNet.Application sap_app = new SaUbiNet.Application();
        public static int count = 0;
        
        static void Main(string[] args)
        {
            SQLConn.ConnectionString = "workstation id=serverzvo;packet size=4096;data source=serverzvo.zvo.LOCAL;timeout=2000;User ID=saperion;Password=saperion;";
            try
            {
                oSAPApp.Login("zagruzka3", "3", 20);
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message);
                return;
            }
            try
            {
                SQLConn.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения к базе данных\n" + ex.Message);
                return;
            }
            IsInclude();
            Console.WriteLine("все, я кончил");
            oSAPApp.Logout();
            SQLConn.Close();
            Console.ReadKey();
        }
        static string IsInclude()
        {
            SqlCommand command2 = new SqlCommand();
            command2.Connection = SQLConn;
            command2.CommandType = CommandType.Text;
            command2.CommandTimeout = 0;
            string xhdoc = "";
            string cardnumber = "";
            command2.CommandText = @"SELECT  [XHDOC], cardnumber      
            FROM  [Saperion].[dbo].[ServCards] where SYSINDEXSTATE=0 and DATEBIRTH is null";

            try
            {
                using (SqlDataReader drr = command2.ExecuteReader())
                {

                    while (drr.Read())
                    {
                        xhdoc = drr["XHDOC"].ToString();
                        cardnumber = drr["cardnumber"].ToString();
                        delete(xhdoc,cardnumber);
                    }
                    drr.Close();
                }
            }
            catch (Exception ex)
            {
                using (FileStream fstream = new FileStream(Application.StartupPath + "\\log.log", FileMode.Append, FileAccess.Write))
                {
                    string log = DateTime.Now.ToString() + " " + ex.Message + "\n";
                    fstream.Write(System.Text.Encoding.Default.GetBytes(log), 0, System.Text.Encoding.Default.GetBytes(log).Length);
                }
                MessageBox.Show("Ошибка обращения к базе данных\n" + ex.Message);
                return null;
            }
            finally
            {
                command2.Dispose();
            }
            return xhdoc;
        }
        public static void delete(string xhdoc,string cardnumber)
        {
            SaUbiNet.Document sap_doc = new SaUbiNet.Document();
            try
            {
                sap_doc.Load(xhdoc);
                SaUbiNet.Cursor sap_cur = sap_app.SelectQuery
                    ("SERVCARDS",
                     "CARDNUMBER='" + sap_doc.GetProperty("CARDNUMBER") + "' AND " +
                     "SYSROWID='" + sap_doc.GetProperty("SYSROWID") + "'");
                sap_cur.DeleteCurrent();
                Console.Clear();
                Console.WriteLine(count++);
            }
            catch (Exception ex)
            {
                using (FileStream fstream = new FileStream(Application.StartupPath + "\\log.log", FileMode.Append, FileAccess.Write))
                {
                    string log = DateTime.Now.ToString() + " не могу удалить " + cardnumber + " "+ ex.Message + "\n";
                    fstream.Write(System.Text.Encoding.Default.GetBytes(log), 0, System.Text.Encoding.Default.GetBytes(log).Length);
                }
            }
            sap_doc.Dispose();
        }
    }
}
