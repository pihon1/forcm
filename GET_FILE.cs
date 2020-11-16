using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using COM;
using GB;

namespace GET_FILE
{
    public partial class GET_FILE : Form
    {
        private string sPath_A = string.Empty;
        private string sPath_B = string.Empty;
        private string sPath_C = string.Empty;
        private string sPath_D = string.Empty;

        private string sLocal_Path = string.Empty;

        public GET_FILE()
        {
            try
            {
                InitializeComponent();
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message);
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Global.gINIPath = Application.StartupPath + "\\INI\\config.ini";
            InitControl();
            //sPath = @"\\192.168.123.196\AutoExport\";
            sPath_A = Global.GetIniValue("BASE SET", "A-LINE", Global.gINIPath);
            sPath_B = Global.GetIniValue("BASE SET", "B-LINE", Global.gINIPath);
            sPath_C = Global.GetIniValue("BASE SET", "C-LINE", Global.gINIPath);
            sPath_D = Global.GetIniValue("BASE SET", "D-LINE", Global.gINIPath);
        }

        private void InitControl()
        {
            #region - SERVER SET
            Global.gDBIP = Global.GetIniValue("SERVER SET", "DBIP", Global.gINIPath);
            Global.gDBNM = Global.GetIniValue("SERVER SET", "DBNM", Global.gINIPath);
            Global.gDBID = Global.GetIniValue("SERVER SET", "DBID", Global.gINIPath);
            Global.gDBPS = Global.GetIniValue("SERVER SET", "DBPS", Global.gINIPath);
            Global.gDBTP = Global.GetIniValue("SERVER SET", "DBTP", Global.gINIPath);

            Global.gVersion = Global.GetIniValue("SERVER SET", "VERSION", Global.gVerPath);

            // 업데이트 프로그랭 경로에 ini파일 생성 후 현재 프로그램의 경로 정보 저장
            if (Global.GetIniValue("Version", "serverpath", Global.gVerPath) == "")
            {
                Global.SetIniValue("Version", "serverpath", @"", Global.gVerPath);
            }

            Global.SetIniValue("Version", "path", Application.StartupPath, Global.gVerPath);
            Global.SetIniValue("Version", "start", Application.ExecutablePath, Global.gVerPath);
            Global.SetIniValue("Version", "app", Application.ProductName, Global.gVerPath);
            #endregion

            #region - BASE SET
            Global.gLanguage = Global.GetIniValue("BASE SET", "LANG", Global.gINIPath);
            Global.gTheme = Global.GetIniValue("BASE SET", "THEME", Global.gINIPath);
            Global.gModuel = Global.GetIniValue("BASE SET", "TPYE", Global.gINIPath);
            Global.gFac = Global.GetIniValue("BASE SET", "FAC", Global.gINIPath);
            Global.gWC = Global.GetIniValue("BASE SET", "WC", Global.gINIPath);
            Global.gCOMPANY = Global.GetIniValue("BASE SET", "COMPANY", Global.gINIPath);

            Global.gLogAt = Global.GetIniValue("LOG SET", "ALOG", Global.gLOGPath);
            #endregion

            //lblWC.Text = Global.gWC + "라인";
        }

        private void SetDirectorySecurity(string linePath)
        {
            //DirectorySecurity dSecurity = Directory.GetAccessControl(linePath);
            //dSecurity.AddAccessRule(new FileSystemAccessRule("Users",
            //                                                            FileSystemRights.FullControl,
            //                                                            InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit,
            //                                                            PropagationFlags.None,
            //                                                            AccessControlType.Allow));
            //Directory.SetAccessControl(linePath, dSecurity);
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            Line_data(sPath_A, "A");
            Line_data(sPath_B, "B");
            Line_data(sPath_C, "C");
            Line_data(sPath_D, "D");
        }

        private void Line_data(string sPath, string sWC)
        {
            try
            {
                string sOrder = string.Empty;
                string sSmt = string.Empty;
                
                FileInfo fileinfo = null;
                
                Dictionary<string, object> dicParam = new Dictionary<string, object>();
                var Files = Directory.EnumerateFiles(sPath, "*.xlsx", SearchOption.TopDirectoryOnly);


               // MessageBox.Show(Files.ToString());
                foreach (string currentFile in Files)
                {
                    // MessageBox.Show(currentFile);
                    try
                    {
                        string s = currentFile; //csv 파일 풀 경로
                        string sDate = DateTime.Now.ToString("yyyyMMddHHmmssfff");

                        DGView_Sheet1.Rows.Count = 0;

                        //생성파일에서 마스터 정보만 추출
                        DataTable dt = Get_Excel_Data(s); //Path.GetFileName(s)



                        //if (DGView_Sheet1.Cells[1, 1].Value.ToString() != "Module")
                        //{
                        //    DGView_Sheet1.DataSource = null;
                        //    return;
                        //}

                        if (dt.Rows[1][1].ToString() != "Module")
                        {
                            dt = null;
                        }
                        else
                        {
                            DGView_Sheet1.DataSource = dt;

                            for (int i = 2; i < DGView.ActiveSheet.RowCount; i++)
                            {
                                dicParam = new Dictionary<string, object>();
                                dicParam.Add("sProcedure", "POP_SCAN_DATA_I100");
                                dicParam.Add("sSection", "SAMSUNG");
                                dicParam.Add("sFac_cd", Global.gFac);
                                dicParam.Add("sWc_cd", Global.gWC);
                                dicParam.Add("sEmp_id", "");
                                dicParam.Add("sV0", DGView_Sheet1.Cells[i, 2].Value.ToString());
                                dicParam.Add("sV1", DGView_Sheet1.Cells[i, 3].Value.ToString());
                                dicParam.Add("sV2", DGView_Sheet1.Cells[i, 4].Value.ToString());
                                dicParam.Add("sV3", DGView_Sheet1.Cells[i, 6].Value.ToString());
                                dicParam.Add("sV4", sWC + "-" + sDate);
                                dicParam.Add("sV5", DGView_Sheet1.Cells[i, 0].Value.ToString());
                                dicParam.Add("sV6", DGView_Sheet1.Cells[i, 8].Value.ToString());
                                dicParam.Add("sV7", DGView_Sheet1.Cells[i, 9].Value.ToString());
                                dicParam.Add("sV8", DGView_Sheet1.Cells[i, 10].Value.ToString());
                                dicParam.Add("sV9", Path.GetFileName(s));

                                sOrder = DGView_Sheet1.Cells[i, 3].Value.ToString();
                                sSmt = DGView_Sheet1.Cells[i, 10].Value.ToString();

                                try
                                {
                                    UseDirect.GetDataSet_N(ControlUtil.BuildConnStr(Global.gDBTP), "COM_PROCEDURE", dicParam);
                                }
                                catch (Exception EX)
                                {
                                    // MessageBox.Show(EX.Message);
                                }
                            }

                        
                            //sLocal_Path = @"D:\SAMSUNG\" + @"\" + sSmt + @"\" + sOrder + @"\";
                            //DirectoryInfo di = new DirectoryInfo(sLocal_Path);
                            //if (di.Exists == false)
                            //{
                            //    di.Create();
                            //}

                            //DB에 INSERT 후 [설비 -> 서버]로 파일이동
                            fileinfo = new FileInfo(s);
                            fileinfo.Delete();
                            //fileinfo.MoveTo(sLocal_Path + Path.GetFileName(s));
                          
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
               // MSGBOX.Show(ex.Message);
            }
        }

        private DataTable Get_Excel_Data(string sFile)
        {
            //MessageBox.Show(sFile);
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + sFile + @";Extended Properties=""Excel 12.0;HDR=NO""";
            string sheetName = string.Empty;

            // 첫 번째 시트의 이름을 가져옮
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }
            Console.WriteLine("sheetName = " + sheetName);

            DataTable dt = new DataTable();
            // 첫 번째 쉬트의 데이타를 읽어서 datagridview 에 보이게 함.
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();
                    }
                }
            }

            return dt;
        }

        int i = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            i++;
            label2.Text = i.ToString() + "/5"; //1분마다 자동 업데이트
            if (i == 5)
            {
                try
                {
                    btnUpdate_Click(null, null);
                    i = 0;
                }
                catch
                {
                    i = 0;
                }
            }
        }



    }
}
