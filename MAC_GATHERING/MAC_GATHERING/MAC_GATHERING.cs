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

namespace MAC_GATHERING
{
    public partial class MAC_GATHERING : Form
    {
        private string sMac_SPI_Path = string.Empty;
        private string sMac_AOI_Path = string.Empty;
        private string sSPI_Path = string.Empty;
        private string sAOI_Path = string.Empty;

        public MAC_GATHERING()
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
            sMac_SPI_Path = Global.GetIniValue("BASE SET", "SPI", Global.gINIPath);
            sMac_AOI_Path = Global.GetIniValue("BASE SET", "AOI", Global.gINIPath);
         
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

            lblWC.Text = Global.gWC + "라인";
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
            try
            {
                string sQR = string.Empty; //마스터QR
                string sDate = string.Empty; //마스터파일 생성일시

                FileInfo fileinfo = null;

                Dictionary<string, object> dicParam = new Dictionary<string, object>();

                var csvFiles = Directory.EnumerateFiles(sMac_SPI_Path, "*.csv", SearchOption.AllDirectories);

                foreach (string currentFile in csvFiles)
                {
                    string s = currentFile; //csv 파일 풀 경로
                    DGView_Sheet1.Rows.Count = 0;
                    DGView2_Sheet1.Rows.Count = 0;

                    if (Path.GetFileName(s).Contains("IDNO")) //파일명에 IDNO 포함된 파일은 제외
                    {
                        fileinfo = new FileInfo(s);
                        fileinfo.MoveTo(sSPI_Path + Path.GetFileName(s));
                    }
                    else
                    {
                        //생성파일에서 마스터 정보만 추출
                        DataTable dt2 = Get_CSV_MASTER(Path.GetFileName(s), "SPI");
                        DGView2_Sheet1.DataSource = dt2;

                        for (int i = 0; i < DGView2.ActiveSheet.RowCount; i++)
                        {
                            //dicParam = new Dictionary<string, object>();
                            //dicParam.Add("sProcedure", "POP_MACHINE_DATA");
                            //dicParam.Add("sSection", "SPI_MASTER");
                            //dicParam.Add("sFac_cd", Global.gFac);
                            //dicParam.Add("sWc_cd", Global.gWC);
                            //dicParam.Add("sEmp_id", "FCM");
                            //dicParam.Add("sV0", i.ToString());
                            //dicParam.Add("sV1", DGView2_Sheet1.Cells[i, 4].Value.ToString()); //QR
                            //dicParam.Add("sV2", Path.GetFileName(s).Substring(0, 4) + "-" + Path.GetFileName(s).Substring(4, 2) + "-" + Path.GetFileName(s).Substring(6, 2));
                            //dicParam.Add("sV3", Path.GetFileName(s).Substring(0, 4) + "-" + Path.GetFileName(s).Substring(4, 2) + "-" + Path.GetFileName(s).Substring(6, 2) + " " + Path.GetFileName(s).Substring(8, 2) + ":" + Path.GetFileName(s).Substring(10, 2) + ":" + Path.GetFileName(s).Substring(12, 2));
                            //dicParam.Add("sV4", DGView2_Sheet1.Cells[i, 6].Value.ToString());//결과
                            //dicParam.Add("sV5", Path.GetFileName(s));
                            //dicParam.Add("sV6", Path.GetFileName(s).Split('_')[1]);

                            //sQR = DGView2_Sheet1.Cells[i, 4].Value.ToString();
                            //sDate = Path.GetFileName(s).Substring(0, 4) + "-" + Path.GetFileName(s).Substring(4, 2) + "-" + Path.GetFileName(s).Substring(6, 2) + " " + Path.GetFileName(s).Substring(8, 2) + ":" + Path.GetFileName(s).Substring(10, 2) + ":" + Path.GetFileName(s).Substring(12, 2);
                
                            //try
                            //{
                            //    UseDirect.GetDataSet_N(ControlUtil.BuildConnStr(Global.gDBTP), "COM_PROCEDURE", dicParam);
                            //}
                            //catch (Exception EX)
                            //{
                            //    MessageBox.Show(EX.Message);
                            //}

                            dicParam = new Dictionary<string, object>();
                            dicParam.Add("sProcedure", "POP_MACHINE_DATA");
                            dicParam.Add("sSection", "SPI_MASTER2");
                            dicParam.Add("sFac_cd", Global.gFac);
                            dicParam.Add("sWc_cd", Global.gWC);
                            dicParam.Add("sEmp_id", DGView2_Sheet1.Cells[i, 7].Value.ToString());
                            dicParam.Add("sV0", Path.GetFileName(s).Substring(0, 4) + "-" + Path.GetFileName(s).Substring(4, 2) + "-" + Path.GetFileName(s).Substring(6, 2));
                            dicParam.Add("sV1", Path.GetFileName(s).Substring(8, 2) + ":" + Path.GetFileName(s).Substring(10, 2) + ":" + Path.GetFileName(s).Substring(12, 2));
                            dicParam.Add("sV2", DGView2_Sheet1.Cells[i, 1].Value.ToString());//MachineID
                            dicParam.Add("sV3", DGView2_Sheet1.Cells[i, 19].Value.ToString());//TotalPanelCnt
                            dicParam.Add("sV4", DGView2_Sheet1.Cells[i, 3].Value.ToString());//TotalArrayCnt
                            dicParam.Add("sV5", DGView2_Sheet1.Cells[i, 18].Value.ToString());//BARCODE
                            dicParam.Add("sV6", DGView2_Sheet1.Cells[i, 5].Value.ToString());//InspectionEndTime
                            dicParam.Add("sV7", DGView2_Sheet1.Cells[i, 6].Value.ToString());//PCBResult
                            dicParam.Add("sV8", DGView2_Sheet1.Cells[i, 7].Value.ToString());//UserID
                            dicParam.Add("sV9", DGView2_Sheet1.Cells[i, 8].Value.ToString());//VolumeMIN
                            dicParam.Add("sV10", DGView2_Sheet1.Cells[i, 9].Value.ToString());//VolumeMAX
                            dicParam.Add("sV11", DGView2_Sheet1.Cells[i, 10].Value.ToString());//HeightMIN
                            dicParam.Add("sV12", DGView2_Sheet1.Cells[i, 11].Value.ToString());//HeightMAX
                            dicParam.Add("sV13", DGView2_Sheet1.Cells[i, 12].Value.ToString());//AreaMIN
                            dicParam.Add("sV14", DGView2_Sheet1.Cells[i, 13].Value.ToString());//AreaMAX
                            dicParam.Add("sV15", DGView2_Sheet1.Cells[i, 14].Value.ToString());//OffsetxMIN
                            dicParam.Add("sV16", DGView2_Sheet1.Cells[i, 15].Value.ToString());//OffsetxMAX
                            dicParam.Add("sV17", DGView2_Sheet1.Cells[i, 16].Value.ToString());//OffsetyMIN
                            dicParam.Add("sV18", DGView2_Sheet1.Cells[i, 17].Value.ToString());//OffsetyMAX
                            dicParam.Add("sV19", Path.GetFileName(s).Substring(15, Path.GetFileName(s).Length - 15).Replace(".CSV", ""));
                            dicParam.Add("sV20", Path.GetFileName(s));

                            try
                            {
                                UseDirect.GetDataSet_N(ControlUtil.BuildConnStr(Global.gDBTP), "COM_PROCEDURE", dicParam);
                            }
                            catch (Exception EX)
                            {
                                MessageBox.Show(EX.Message);
                            }

                        }



                        /////////////////////////////////////////////////////////////////////////////////////
                        //생성파일에서 상세검사 정보만 추출
                        DataTable dt = Get_CSV_DATA(Path.GetFileName(s), "SPI");
                        DGView_Sheet1.DataSource = dt;

                        //for (int i = 1; i < DGView.ActiveSheet.RowCount; i++)
                        //{

                        //    dicParam = new Dictionary<string, object>();
                        //    dicParam.Add("sProcedure", "POP_MACHINE_DATA");
                        //    dicParam.Add("sSection", "SPI_DATA");
                        //    dicParam.Add("sFac_cd", "");
                        //    dicParam.Add("sWc_cd", "MES");
                        //    dicParam.Add("sEmp_id", "FCM");
                        //    dicParam.Add("sV0", sDate);
                        //    dicParam.Add("sV1", sQR); //QRA
                        //    dicParam.Add("sV2", DGView_Sheet1.Cells[i, 0].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 1].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 2].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 3].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 4].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 5].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 6].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 7].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 8].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 9].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 10].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 11].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 12].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 13].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 14].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 15].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 16].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 17].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 18].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 19].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 20].Value.ToString() + "\t" +
                        //                        DGView_Sheet1.Cells[i, 21].Value.ToString());
                        //    dicParam.Add("sV3", DGView_Sheet1.Cells[i, 1].Value.ToString()); //PANEL NO

                        //    try
                        //    {
                        //        UseDirect.GetDataSet_N(ControlUtil.BuildConnStr(Global.gDBTP), "COM_PROCEDURE", dicParam);
                        //    }
                        //    catch (Exception EX)
                        //    {
                        //        MessageBox.Show(EX.Message);
                        //    }
                        //}

                        try
                        {
                            string date = Path.GetFileName(s).Substring(0, 4) + "-" + Path.GetFileName(s).Substring(4, 2) + "-" + Path.GetFileName(s).Substring(6, 2) + " " + Path.GetFileName(s).Substring(8, 2) + ":" + Path.GetFileName(s).Substring(10, 2) + ":" + Path.GetFileName(s).Substring(12, 2);

                            sSPI_Path = @"D:\SPI\" + lblWC.Text + @"\" + Convert.ToDateTime(date).ToString("yyyy-MM") + @"\" + Convert.ToDateTime(date).ToString("yyyyMMdd") + @"\";
                            DirectoryInfo di = new DirectoryInfo(sSPI_Path);
                            if (di.Exists == false)
                            {
                                di.Create();
                            }

                            //DB에 INSERT 후 [설비 -> 서버]로 파일이동
                            fileinfo = new FileInfo(s);
                            fileinfo.MoveTo(sSPI_Path + Path.GetFileName(s));
                        }
                        catch { }
                        ////////////////////////////////////////////////////////////////////////////////////////////////////

                    }
                }
            }
            catch(Exception ex)
            {
                
               // MSGBOX.Show(ex.Message);
            }
           
        }


        private void btnAOI_Click(object sender, EventArgs e)
        {
            try
            {
                //sAOI_Path = @"D:\AOI\" + lblWC.Text + @"\" + DateTime.Now.ToString("yyyy-MM") + @"\" + DateTime.Now.ToString("yyyyMMdd") + @"\";
                //DirectoryInfo di = new DirectoryInfo(sAOI_Path);
                //if (di.Exists == false)
                //{
                //    di.Create();
                //}

                string sQR = string.Empty;
                string AOI_Dir = string.Empty;

                FileInfo fileinfo = null;

                Dictionary<string, object> dicParam = new Dictionary<string, object>();

                var csvFiles = Directory.EnumerateFiles(sMac_AOI_Path, "*.csv", SearchOption.AllDirectories);

                foreach (string currentFile in csvFiles)
                {
                    string s = currentFile; //csv 파일 풀 경로
                    DGView_AOI_Sheet1.Rows.Count = 0;

                    DataTable dt = Get_CSV_MASTER(Path.GetFileName(s),"AOI");
                    DGView_AOI_Sheet1.DataSource = dt;

                    //string Model = Path.GetFileName(s).Replace(Path.GetFileName(s).Substring(Path.GetFileName(s).Length - 22,22),"").Replace("AOI-XL033_","");
                    //string date = Path.GetFileName(s).Substring(Path.GetFileName(s).Length - 21, 14);
                    string Model = Path.GetFileName(s).Substring(0,Path.GetFileName(s).Length - 22);
                    string date = Path.GetFileName(s).Substring(Path.GetFileName(s).Length - 21, 14);
                    string Side = "TOP";

                    
                    if (Model.Contains("TOP"))
                        Side = "TOP";
                    else if (Model.Contains("BOT"))
                        Side = "BOT";

                    for (int i = 0; i < DGView_AOI.ActiveSheet.RowCount; i++)
                    {
                        dicParam = new Dictionary<string, object>();
                        dicParam.Add("sProcedure", "POP_MACHINE_DATA");
                        dicParam.Add("sSection", "AOI_MASTER");
                        dicParam.Add("sFac_cd", Global.gFac);
                        dicParam.Add("sWc_cd", Global.gWC);
                        dicParam.Add("sEmp_id", "FCM");
                        dicParam.Add("sV0", Model);
                        dicParam.Add("sV1", DGView_AOI_Sheet1.Cells[i, 1].Value.ToString()); //QR
                        dicParam.Add("sV2", DGView_AOI_Sheet1.Cells[i, 2].Value.ToString()); //결과 ok,ng
                        dicParam.Add("sV3", date.Substring(0, 4) + "-" + date.Substring(4, 2) + "-" + date.Substring(6, 2));
                        dicParam.Add("sV4", date.Substring(0, 4) + "-" + date.Substring(4, 2) + "-" + date.Substring(6, 2) + " " + date.Substring(8, 2) + ":" + date.Substring(10, 2) + ":" + date.Substring(12, 2));
                        dicParam.Add("sV5", Path.GetFileName(s));
                        dicParam.Add("sV6", DGView_AOI_Sheet1.Cells[i, 3].Value.ToString()); //검사시작일시
                        dicParam.Add("sV7", DGView_AOI_Sheet1.Cells[i, 4].Value.ToString()); //검사종료일시
                        dicParam.Add("sV8", Side); //작업면

                        try
                        {
                            UseDirect.GetDataSet_N(ControlUtil.BuildConnStr(Global.gDBTP), "COM_PROCEDURE", dicParam);
                        }
                        catch (Exception EX)
                        {
                            MessageBox.Show(EX.Message);
                        }

                    }


                    //파일이동
                    try
                    {

                        sAOI_Path = @"D:\AOI\" + lblWC.Text + @"\" + Convert.ToDateTime(date.Substring(0, 4) + "-" + date.Substring(4, 2) + "-" + date.Substring(6, 2) + " " + date.Substring(8, 2) + ":" + date.Substring(10, 2) + ":" + date.Substring(12, 2)).ToString("yyyy-MM") + @"\" + Convert.ToDateTime(date.Substring(0, 4) + "-" + date.Substring(4, 2) + "-" + date.Substring(6, 2) + " " + date.Substring(8, 2) + ":" + date.Substring(10, 2) + ":" + date.Substring(12, 2)).ToString("yyyyMMdd") + @"\";
                        DirectoryInfo di = new DirectoryInfo(sAOI_Path);
                        if (di.Exists == false)
                        {
                            di.Create();
                        }

                        fileinfo = new FileInfo(s);
                        fileinfo.MoveTo(sAOI_Path + Path.GetFileName(s));
                    }
                    catch { }
                    /////////////////////////////////////////////////////////////////////////////////////
                }
            }
            catch (Exception ex)
            {
               // MSGBOX.Show(ex.Message);
            }
           
        }

        private DataTable Get_CSV_MASTER(string sFile, string gb)
        {

            if (gb == "SPI")
            {
                DataTable _dt = new DataTable();
                // 컬럼명과 컬럼헤더를 사용해 컬럼을 정의한다
                _dt.Columns.Add("JobName");
                _dt.Columns.Add("MachineID");
                _dt.Columns.Add("TotalPanelCnt");
                _dt.Columns.Add("TotalArrayCnt");
                _dt.Columns.Add("MasterBarcode");
                _dt.Columns.Add("InspectionEndTime");
                _dt.Columns.Add("PCBResult");
                _dt.Columns.Add("UserID");
                _dt.Columns.Add("VolumeMIN");
                _dt.Columns.Add("VolumeMAX");
                _dt.Columns.Add("HeightMIN");
                _dt.Columns.Add("HeightMAX");
                _dt.Columns.Add("AreaMIN");
                _dt.Columns.Add("AreaMAX");
                _dt.Columns.Add("OffsetxMIN");
                _dt.Columns.Add("OffsetxMAX");
                _dt.Columns.Add("OffsetyMIN");
                _dt.Columns.Add("OffsetyMAX");
                _dt.Columns.Add("BARCODE");
                _dt.Columns.Add("PANEL_NO");

                // 데이타를 읽는 StreamReader
                StreamReader rd = new StreamReader(sMac_SPI_Path + "\\" + sFile);
                string[] cols_main = null;
                // 마지막이 될 때까지 루프
                while (!rd.EndOfStream)
                {
                    // 한 라인을 읽는다
                    string line = rd.ReadLine();
                    
                    // 라인을 콤마로 분리하여 컬럼을 만든다
                    string[] cols = line.Split(',');

                    if (cols.Length == 18)
                    {
                        cols_main = line.Split(',');
                        // 한 라인에 각 컬럼의 데이타를 순서대로 넣는다
                        //_dt.Rows.Add(cols[0], cols[1], cols[2], cols[3], cols[4], cols[5],
                        //             cols[6], cols[7], cols[8], cols[9], cols[10], cols[11],
                        //             cols[12], cols[13], cols[14], cols[15], cols[16], cols[17]);
                    }
                    else
                    {
                        if (cols.Length != 1)
                        {
                            if (cols[0] == "ComponentID")//ComponentID
                            {
                                break;
                            }

                            _dt.Rows.Add(cols_main[0], cols_main[1], cols_main[2], cols_main[3], cols_main[4], cols_main[5],
                                        cols_main[6], cols_main[7], cols_main[8], cols_main[9], cols_main[10], cols_main[11],
                                        cols_main[12], cols_main[13], cols_main[14], cols_main[15], cols_main[16], cols_main[17],
                                        cols[0], cols[1]);
                        }
                    }

                    //if (_dt.Rows.Count != 0)
                    //{
                    //    if (_dt.Rows[_dt.Rows.Count - 1][18].ToString() == "ComponentID")//ComponentID
                    //    {
                    //        break;
                    //    }
                    //}
                }

                // StreamReader는 사용 후 반드시 닫는다
                rd.Close();
                _dt.Rows[0].Delete();
                _dt.AcceptChanges();
                return _dt;
            }
            else
            {
                DataTable _dt = new DataTable();
                // 컬럼명과 컬럼헤더를 사용해 컬럼을 정의한다
                _dt.Columns.Add("SEQ");
                _dt.Columns.Add("QR_CODE");
                _dt.Columns.Add("RESULT");
                _dt.Columns.Add("START_DATE");
                _dt.Columns.Add("END_DATE");

                // 데이타를 읽는 StreamReader
                StreamReader rd = new StreamReader(sMac_AOI_Path + "\\" + sFile);

                // 마지막이 될 때까지 루프
                while (!rd.EndOfStream)
                {
                    // 한 라인을 읽는다
                    string line = rd.ReadLine();

                    // 라인을 콤마로 분리하여 컬럼을 만든다
                    string[] cols = line.Split(',');

                    // 한 라인에 각 컬럼의 데이타를 순서대로 넣는다
                    _dt.Rows.Add(cols[0], cols[1], cols[2], cols[3], cols[4]);
                }

                // StreamReader는 사용 후 반드시 닫는다
                rd.Close();
                return _dt;
            }

           
        }

        private DataTable Get_CSV_DATA(string sFile, string gb)
        {
            //string Sql = @"SELECT * FROM [" + sFile + "]";

            //using (OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sPath + "; Extended Properties=\"Text;HDR=No;IMEX=1\""))
            //using (OleDbCommand comm = new OleDbCommand(Sql, conn))
            //using (OleDbDataAdapter adp = new OleDbDataAdapter(comm))
            //{
            //    DataTable _dt = new DataTable();
            //    _dt.Locale = CultureInfo.CurrentCulture;
            //    adp.Fill(_dt);
            //    _dt.Rows[0].Delete();
            //    _dt.Rows[1].Delete();
            //    _dt.AcceptChanges();
            //    return _dt;
            //}

            DataTable _dt = new DataTable();
            // 컬럼명과 컬럼헤더를 사용해 컬럼을 정의한다
            _dt.Columns.Add("ComponentID");
            _dt.Columns.Add("PanelNumber");
            _dt.Columns.Add("ArrayNumber");
            _dt.Columns.Add("PinNumber");
            _dt.Columns.Add("PadDefectType");
            _dt.Columns.Add("Volume");
            _dt.Columns.Add("Area");
            _dt.Columns.Add("Height");
            _dt.Columns.Add("OffsetX");
            _dt.Columns.Add("OffsetY");
            _dt.Columns.Add("PositionX");
            _dt.Columns.Add("PositionY");
            _dt.Columns.Add("VolumeUpper");
            _dt.Columns.Add("VolumeLower");
            _dt.Columns.Add("AreaUpper");
            _dt.Columns.Add("AreaLower");
            _dt.Columns.Add("HeightUpper");
            _dt.Columns.Add("HeightLower");
            _dt.Columns.Add("OffsetXUpper");
            _dt.Columns.Add("OffsetXLower");
            _dt.Columns.Add("OffsetYUpper");
            _dt.Columns.Add("OffsetYLower");

            // 데이타를 읽는 StreamReader
            StreamReader rd = new StreamReader(sMac_SPI_Path + "\\" + sFile);

            // 마지막이 될 때까지 루프
            while (!rd.EndOfStream)
            {
                // 한 라인을 읽는다
                string line = rd.ReadLine();

                // 라인을 콤마로 분리하여 컬럼을 만든다
                int Len = line.Split(',').Length;

                // 총 컬럼 수 = 22
                // 컬럼 수가 21일 경우 ',' 붙여서 강제로 22로 맞춤
                if (Len == 21)
                {
                    line += ",";
                }
                string[] cols = line.Split(',');

                if (cols.Length == 22)
                {
                    // 한 라인에 각 컬럼의 데이타를 순서대로 넣는다
                    _dt.Rows.Add(cols[0], cols[1], cols[2], cols[3], cols[4], cols[5],
                                 cols[6], cols[7], cols[8], cols[9], cols[10], cols[11],
                                 cols[12], cols[13], cols[14], cols[15], cols[16], cols[17],
                                 cols[18], cols[19], cols[20], cols[21]);
                }
            }

            // StreamReader는 사용 후 반드시 닫는다
            rd.Close();
            _dt.Rows[0].Delete();
            _dt.AcceptChanges();
            return _dt;
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
                    btnAOI_Click(null, null);
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
