//, row["cell_type"].ToString()
//               , row["not_null"].ToString()
//               , row["yn_sort"].ToString()
//               , row["yn_filter"].ToString()


// if (WORK_DAY_KEY1.Value > WORK_DAY_KEY2.Value)
//            {
//                MSGBOX.Show("From일자가 To일자를 초과할 수 없습니다.");
//                WORK_DAY_KEY1.Focus();
//                return;
//            }

//// 그리드 변경 내용 있는지 체크
//if (!modFPGrid.Change_Grid(DGView_Sheet1))
//    return;

//// 그리드 필수값 체크
//if (!modFPGrid.Required_Grid(DGView_Sheet1))
//    return;

//// 중복값 체크
//for (int i = 0; i < DGView_Sheet1.Rows.Count; i++)
//{
//    if (DGView_Sheet1.Rows[i].Tag == null) DGView_Sheet1.Rows[i].Tag = "";

//    if (DGView_Sheet1.Rows[i].Tag.ToString() == "INSERT")
//    {
//        DataSet ds = ControlUtil.MST_CHECK("MST_USER", DGView_Sheet1.Cells[i, 1].Value.ToString());
//        if (ds.Tables[0].Rows.Count != 0)
//        {
//            MSGBOX.Show("이미 등록된 사번입니다.\n중복값 : " + ds.Tables[0].Rows[0][0].ToString());
//            ds.Clear();
//            return;
//        }
//    }
//}
//MSGBOX.Show("저장되었습니다.", "알림", 15);

//if (DGView_Sheet1.Rows.Count == 0)
//            {
//                MSGBOX.Show("삭제할 데이터가 없습니다.");
//                return;
//            }

// //해당 필드에 입력가능한 최대 값 체크
//            if(!modFPGrid.Max_Len_Col("POP_MAT_17_01", DGView_Sheet1))
//                return;


//if (DGView_Sheet1.RowCount > 0)
//            {
//                GB.modFile.FarpointToExcel(this.DGView_Sheet1);
//            }
//            else
//            {
//                MSGBOX.Show("화면에 데이터가 존재하지 않아 사용 불가합니다.");
//            }

//////해당 필드에 입력가능한 최대 값 체크(0일 경우 체크 안함)
////int Max_len = modFPGrid.Max_Len_Col("MES_BASIC_01_2", e.Column);
////if (Max_len != 0 && DGView_Sheet1.Cells[e.Row, e.Column].Value != null)
////{
////    if (DGView_Sheet1.Cells[e.Row, e.Column].Value.ToString().Length > Max_len)
////    {
////        MSGBOX.Show("최대 입력 자릿수를 초과했습니다.\n(최대자릿수:" + Max_len + ")");
////        if (DGView_Sheet1.Cells[e.Row, e.Column].Value.ToString().Length > Max_len)
////            DGView_Sheet1.Cells[e.Row, e.Column].Value = DGView_Sheet1.Cells[e.Row, e.Column].Value.ToString().Substring(0, Max_len);
////    }
////}