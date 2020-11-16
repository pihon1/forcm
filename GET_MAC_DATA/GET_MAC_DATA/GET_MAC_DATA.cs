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
    public partial class GET_MAC_DATA : Form
    {
        private string sPath_SPI_A = string.Empty;
        private string sPath_SPI_B = string.Empty;
        private string sPath_SPI_C = string.Empty;
        private string sPath_SPI_D = string.Empty;

        private string sPath_AOI_A = string.Empty;
        private string sPath_AOI_B = string.Empty;
        private string sPath_AOI_C = string.Empty;
        private string sPath_AOI_D = string.Empty;

        private string sLocal_Path = string.Empty;

        private string sModule = string.Empty;

        public GET_MAC_DATA()
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
            sPath_SPI_A = Global.GetIniValue("BASE SET", "A-LINE-SPI", Global.gINIPath);
            sPath_SPI_B = Global.GetIniValue("BASE SET", "B-LINE-SPI", Global.gINIPath);
            sPath_SPI_C = Global.GetIniValue("BASE SET", "C-LINE-SPI", Global.gINIPath);
            sPath_SPI_D = Global.GetIniValue("BASE SET", "D-LINE-SPI", Global.gINIPath);

            sPath_AOI_A = Global.GetIniValue("BASE SET", "A-LINE-AOI", Global.gINIPath);
            sPath_AOI_B = Global.GetIniValue("BASE SET", "B-LINE-AOI", Global.gINIPath);
            sPath_AOI_C = Global.GetIniValue("BASE SET", "C-LINE-AOI", Global.gINIPath);
            sPath_AOI_D = Global.GetIniValue("BASE SET", "D-LINE-AOI", Global.gINIPath);
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
            Line_data(sPath_SPI_A, "A", "SPI");
            Line_data(sPath_SPI_B, "B", "SPI");
            Line_data(sPath_SPI_C, "C", "SPI");
            Line_data(sPath_SPI_D, "D", "SPI");

            Line_data(sPath_AOI_A, "A", "AOI");
            Line_data(sPath_AOI_B, "B", "AOI");
            Line_data(sPath_AOI_C, "C", "AOI");
            Line_data(sPath_AOI_D, "D", "AOI");
        }

        private void Line_data(string sPath, string sWC, string sMac)
        {
            try
            {
                string sOrder = string.Empty;
                string sSmt = string.Empty;
                
                FileInfo fileinfo = null;

                Dictionary<string, object> dicParam = new Dictionary<string, object>();

                var Files = Directory.EnumerateFiles(sPath, "*.csv", SearchOption.AllDirectories);

                foreach (string currentFile in Files)
                {
                    string s = currentFile; //csv 파일 풀 경로
                    DGView_Sheet1.Rows.Count = 0;

                    //생성파일에서 마스터 정보만 추출
                    DataTable dt = Get_Excel_Data(s, sMac); //Path.GetFileName(s)

                    DGView_Sheet1.DataSource = dt;

                    //if (DGView_Sheet1.Cells[1, 1].Value.ToString() != "Module")
                    //{
                    //    return;
                    //}

                    for ( int i = 1; i < DGView.ActiveSheet.RowCount; i++)
                    {
                        sModule = DGView_Sheet1.Cells[i, 3].Value.ToString();

                        dicParam = new Dictionary<string, object>();
                        dicParam.Add("sProcedure", "POP_WIP_MAC_DATA");
                        dicParam.Add("sSection", sMac);
                        dicParam.Add("sFac_cd", Global.gFac);
                        dicParam.Add("sWc_cd", sWC);
                        dicParam.Add("sEmp_id", "");
                        dicParam.Add("sV0", DGView_Sheet1.Cells[i, 0].Value.ToString());
                        dicParam.Add("sV1", DGView_Sheet1.Cells[i, 1].Value.ToString());
                        dicParam.Add("sV2", DGView_Sheet1.Cells[i, 2].Value.ToString());
                        dicParam.Add("sV3", DGView_Sheet1.Cells[i, 3].Value.ToString());
                        dicParam.Add("sV4", DGView_Sheet1.Cells[i, 4].Value.ToString());
                        dicParam.Add("sV5", DGView_Sheet1.Cells[i, 5].Value.ToString());
                        dicParam.Add("sV6", DGView_Sheet1.Cells[i, 6].Value.ToString());
                        dicParam.Add("sV7", DGView_Sheet1.Cells[i, 7].Value.ToString());
                        dicParam.Add("sV8", DGView_Sheet1.Cells[i, 8].Value.ToString());
                        dicParam.Add("sV9", DGView_Sheet1.Cells[i, 9].Value.ToString());

                        if (sMac == "SPI")
                        {
                            dicParam.Add("sV10", DGView_Sheet1.Cells[i, 10].Value.ToString());
                            dicParam.Add("sV11", DGView_Sheet1.Cells[i, 11].Value.ToString());
                            dicParam.Add("sV12", DGView_Sheet1.Cells[i, 12].Value.ToString());
                            dicParam.Add("sV13", DGView_Sheet1.Cells[i, 13].Value.ToString());
                            dicParam.Add("sV14", DGView_Sheet1.Cells[i, 14].Value.ToString());
                            dicParam.Add("sV15", DGView_Sheet1.Cells[i, 15].Value.ToString());
                            dicParam.Add("sV16", DGView_Sheet1.Cells[i, 16].Value.ToString());
                            dicParam.Add("sV17", DGView_Sheet1.Cells[i, 17].Value.ToString());
                            dicParam.Add("sV18", DGView_Sheet1.Cells[i, 18].Value.ToString());
                            dicParam.Add("sV19", DGView_Sheet1.Cells[i, 19].Value.ToString());
                            dicParam.Add("sV20", DGView_Sheet1.Cells[i, 20].Value.ToString());
                            dicParam.Add("sV21", DGView_Sheet1.Cells[i, 21].Value.ToString());
                            dicParam.Add("sV22", DGView_Sheet1.Cells[i, 22].Value.ToString());
                            dicParam.Add("sV23", DGView_Sheet1.Cells[i, 23].Value.ToString());
                            dicParam.Add("sV24", DGView_Sheet1.Cells[i, 24].Value.ToString());
                            dicParam.Add("sV25", DGView_Sheet1.Cells[i, 25].Value.ToString());
                            dicParam.Add("sV26", DGView_Sheet1.Cells[i, 26].Value.ToString());
                            dicParam.Add("sV27", DGView_Sheet1.Cells[i, 27].Value.ToString());
                            dicParam.Add("sV28", DGView_Sheet1.Cells[i, 28].Value.ToString());
                            dicParam.Add("sV29", DGView_Sheet1.Cells[i, 29].Value.ToString());
                            dicParam.Add("sV30", DGView_Sheet1.Cells[i, 30].Value.ToString());
                            dicParam.Add("sV31", DGView_Sheet1.Cells[i, 31].Value.ToString());
                            dicParam.Add("sV32", DGView_Sheet1.Cells[i, 32].Value.ToString());
                        }

                        dicParam.Add("sV33", Path.GetFileName(s));

                        try
                        {
                            UseDirect.GetDataSet_N(ControlUtil.BuildConnStr(Global.gDBTP), "COM_PROCEDURE", dicParam);
                        }
                        catch (Exception EX)
                        {
                            MessageBox.Show(EX.Message);
                        }
                    }

                    try
                    {

                        sLocal_Path = @"D:\" + sMac + @"\" + @"\" + sWC + @"라인\" + DateTime.Now.ToString("yyyyMMdd") + @"\";
                        DirectoryInfo di = new DirectoryInfo(sLocal_Path);
                        if (di.Exists == false)
                        {
                            di.Create();
                        }

                        //DB에 INSERT 후 [설비 -> 서버]로 파일이동
                        fileinfo = new FileInfo(s);
                        //fileinfo.Delete();
                        fileinfo.MoveTo(sLocal_Path + Path.GetFileName(s));

                        
                    }
                    catch { }

                }
            }
            catch (Exception ex)
            {

                // MSGBOX.Show(ex.Message);
            }
        }

        private DataTable Get_Excel_Data(string path, string mac)
        {
            int length = 0;
            if (mac == "SPI")
                length = 32;
            else
                length = 11;

            DataTable _dt = new DataTable();
            // 컬럼명과 컬럼헤더를 사용해 컬럼을 정의한다
            _dt.Columns.Add("C1");
            _dt.Columns.Add("C2");
            _dt.Columns.Add("C3");
            _dt.Columns.Add("C4");
            _dt.Columns.Add("C5");
            _dt.Columns.Add("C6");
            _dt.Columns.Add("C7");
            _dt.Columns.Add("C8");
            _dt.Columns.Add("C9");
            _dt.Columns.Add("C10");

            if (mac == "SPI")
            {
                _dt.Columns.Add("C11");
                _dt.Columns.Add("C12");
                _dt.Columns.Add("C13");
                _dt.Columns.Add("C14");
                _dt.Columns.Add("C15");
                _dt.Columns.Add("C16");
                _dt.Columns.Add("C17");
                _dt.Columns.Add("C18");
                _dt.Columns.Add("C19");
                _dt.Columns.Add("C20");
                _dt.Columns.Add("C21");
                _dt.Columns.Add("C22");
                _dt.Columns.Add("C23");
                _dt.Columns.Add("C24");
                _dt.Columns.Add("C25");
                _dt.Columns.Add("C26");
                _dt.Columns.Add("C27");
                _dt.Columns.Add("C28");
                _dt.Columns.Add("C29");
                _dt.Columns.Add("C30");
                _dt.Columns.Add("C31");
                _dt.Columns.Add("C32");
                _dt.Columns.Add("C33");
            }

            // 데이타를 읽는 StreamReader
            StreamReader rd = new StreamReader(path);

            // 마지막이 될 때까지 루프
            while (!rd.EndOfStream)
            {
                // 한 라인을 읽는다
                string line = rd.ReadLine();

                // 라인을 콤마로 분리하여 컬럼을 만든다
                int Len = line.Split(',').Length;

                // 총 컬럼 수 = 22
                // 컬럼 수가 21일 경우 ',' 붙여서 강제로 22로 맞춤
                if (Len == length)
                {
                    line += ",";
                }
                string[] cols = line.Split(',');

                if (cols.Length == length + 1)
                {
                    // 한 라인에 각 컬럼의 데이타를 순서대로 넣는다

                    if (mac == "SPI")
                    {
                        _dt.Rows.Add(cols[0], cols[1], cols[2], cols[3], cols[4], cols[5],
                                     cols[6], cols[7], cols[8], cols[9], cols[10], cols[11],
                                     cols[12], cols[13], cols[14], cols[15], cols[16], cols[17],
                                     cols[18], cols[19], cols[20], cols[21], cols[22], cols[23],
                                     cols[24], cols[25], cols[26], cols[27], cols[28], cols[29],
                                     cols[30], cols[31], cols[32]);
                    }
                    else
                    {
                        _dt.Rows.Add(cols[0], cols[1], cols[2], cols[3], cols[4], cols[5],
                                     cols[6], cols[7], cols[8], cols[9]);
                    }
                }
            }

            // StreamReader는 사용 후 반드시 닫는다
            rd.Close();
            _dt.Rows[0].Delete();
            _dt.AcceptChanges();
            return _dt;
        }

        private void Set_Barcode(string sWC, string sMac)
        {
            Dictionary<string, object> dicParam = new Dictionary<string, object>();
            dicParam = new Dictionary<string, object>();
            dicParam.Add("sProcedure", "POP_WIP_MAC_MODULE");
            dicParam.Add("sSection", sMac);
            dicParam.Add("sFac_cd", Global.gFac);
            dicParam.Add("sWc_cd", sWC);
            dicParam.Add("sEmp_id", "");
            dicParam.Add("sV0", sModule);

            try
            {
                UseDirect.GetDataSet_N(ControlUtil.BuildConnStr(Global.gDBTP), "COM_PROCEDURE", dicParam);
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.Message);
            }
        }

        int i = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            i++;
            label2.Text = i.ToString() + "/60"; //1분마다 자동 업데이트
            if (i == 60)
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

        private void button1_Click(object sender, EventArgs e)
        {

        }



    }
}
