#region < HEADER AREA >
/*--------------------------------------------------------------------------------------------
CREATE       : 2021-06-09 JEJUN  
DESCRIPT    :  엑셀리포트처리 - 헬퍼에서 처리된 데이터를 받아와 실제 처리 구현.

UPDATE       :
DESCRIPT    :  
---------------------------------------------------------------------------------------------*/
#endregion

#region < USING AREA >
using System;
using System.Data;
using System.Threading;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Drawing;
using System.IO;
#endregion

namespace ExcelReportHelper
{
    internal class ExcelReport
    {
        //멤버변수
        Microsoft.Office.Interop.Excel.Application excelApp = null;
        Microsoft.Office.Interop.Excel.Workbook wb = null;
        Microsoft.Office.Interop.Excel.Worksheet ws1 = null; //sheet1

        string sFileName = "";
        DataSet ds = null;
        DataTable dt = null;
        int[] m_CellLoc;

        List<ExeclReportDtInfo> dtinfo;
        ExeclReportDtInfo m_dtinfo;

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sFileName"> 출력 파일명 (경로포함 fullname) </param>
        /// <param name="dtinfo"> 데이터그룹정보 </param>
        /// <param name="dataSet"> 출력 데이터셋 </param>
        internal ExcelReport(string sFileName, List<ExeclReportDtInfo> execlReportDtInfo, DataSet dataSet)
        {
            this.sFileName = sFileName;
            this.dtinfo = execlReportDtInfo;
            this.ds = dataSet;
        }

        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        // 셀을 x,y 값으로
        private static int[] CellToXY(string s)
        {
            int[] returnobj = new int[] { 1, 1, 1, 1 };
            string sX = "1";
            string sY = "A";
            string[] sR = s.Split(':');
            // x 숫자짜르기
            sX = Regex.Replace(sR[0], @"\D", "");//숫자 추출
            returnobj[0] = string.IsNullOrEmpty(sX) ? 1 : Convert.ToInt32(sX);
            returnobj[2] = string.IsNullOrEmpty(sX) ? 1 : Convert.ToInt32(sX);
            // y 
            sY = Regex.Replace(sR[0], @"\d", "");//문자 추출
            returnobj[1] = iTransAlpha(sY);
            returnobj[3] = iTransAlpha(sY);

            if (sR.Length > 1)
            {
                sX = Regex.Replace(sR[1], @"\D", "");//숫자 추출
                returnobj[2] = string.IsNullOrEmpty(sX) ? 1 : Convert.ToInt32(sX);
                // y 
                sY = Regex.Replace(sR[1], @"\d", "");//문자 추출
                returnobj[3] = iTransAlpha(sY);
            }

            return returnobj;
        }

        private static int iTransAlpha(string s)
        {
            int iValue = 0;

            if (s.Length > 0)
            {
                char c = s[s.Length - 1];
                s = s.Substring(0, s.Length - 1);

                iValue = ((char)c - 'A') + 1;

                if (s.Length > 0)
                    iValue = iValue + 26 * iTransAlpha(s);
            }
            return iValue;
        }

        /// <summary>
        /// 해당 시트 범위에서 캡션으로 위치 찾기     
        /// 행 기준 찾기 , 열 기준 찾기 , 전체찾기 
        /// </summary>
        /// <param name="array">시트범위 </param>
        /// <param name="elem">찾을 캡션 </param>
        /// <param name="orientation"> 방향 </param>
        /// <param name="home">시작위치 </param>
        /// <returns></returns>
        static int[] FindCell(Microsoft.Office.Interop.Excel.Worksheet ws, string elem, string orientation, string range)
        {
            Microsoft.Office.Interop.Excel.Range rng = ws.UsedRange;   //현재 시트에서 사용중인 범위
            object[,] array = rng.Value; //범위의 데이터

            int[] returnVal = new int[4] { 0, 0, 0, 0 }; // start row , start column, end row, end solumn
            int rowEndIndx = array.GetLength(0);
            int colEndIndx = array.GetLength(1);

            int[] homeindex = CellToXY(range);
            int rowStartIndx = homeindex[0];
            int colStartIndx = homeindex[1];

            if (orientation.Equals("N"))
                return homeindex;

            //if (orientation.Equals( "V"))
            //    colEndIndx = colStartIndx;
            //if (orientation.Equals("H"))
            //    rowEndIndx = rowStartIndx;
            //else if (orientation.Equals("N"))
            //    return homeindex;

            for (int rowIndex = rowStartIndx; rowIndex <= rowEndIndx; rowIndex++)
            {
                for (int colIndex = colStartIndx; colIndex <= colEndIndx; colIndex++)
                {
                    if (Convert.ToString(array[rowIndex, colIndex]).Replace(" ", string.Empty) == elem) // find caption 
                    {
                        // find merge 
                        Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)ws.Cells[rowIndex, colIndex];
                        Microsoft.Office.Interop.Excel.Range oRange2 = (Microsoft.Office.Interop.Excel.Range)oRange.MergeArea;
                        int iLastCol = colIndex + oRange2.Columns.Count - 1;
                        int iLastRow = rowIndex + oRange2.Rows.Count - 1;

                        returnVal.SetValue(rowIndex, 0);
                        returnVal.SetValue(colIndex, 1);
                        returnVal.SetValue(iLastRow, 2);
                        returnVal.SetValue(iLastCol, 3);
                        return returnVal;
                    }
                }
            }
            return returnVal;
        }

        /// <summary>
        /// 프린트 
        /// </summary>
        /// <param name="sPrintName"> 선택프린터이름</param>
        internal void PrintReport(string sPrintName, bool bSave)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range oRange = null;
                Microsoft.Office.Interop.Excel.Range oRange2 = null;

                //-------------------------------
                // 엑셀 오픈                   
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                wb = excelApp.Workbooks.Open(this.sFileName);
                ws1 = wb.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet;
                //Microsoft.Office.Interop.Excel.Range rng = ws1.UsedRange;   //현재 시트에서 사용중인 범위
                //rngData = rng.Value; //범위의 데이터

                //-------------------------------                
                int iRepeatCnt = 0;  // 테이블 반복 횟수 
                for (int iLoop1 = 0; iLoop1 < dtinfo.Count; iLoop1++) // 데이터테이블
                {
                    iRepeatCnt = 0;
                    dt = ds.Tables[iLoop1];
                    m_dtinfo = dtinfo[iLoop1];
                    for (int iLoop2 = 0; iLoop2 < dt.Rows.Count; iLoop2++) // 로우
                    {
                        // MAXROW 넘었을때.
                        if (iLoop2 >= m_dtinfo.iMaxRow)
                        {
                            switch (m_dtinfo.sContinueMode)
                            {
                                case "A":      // 로우 추가      
                                    if (m_dtinfo.iH.Equals((int)1) ? true : false)
                                    {
                                        oRange = (Microsoft.Office.Interop.Excel.Range)ws1.Cells[m_CellLoc[2] + iLoop2, 1];
                                        oRange = oRange.EntireRow;
                                        oRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                                        oRange2 = (Microsoft.Office.Interop.Excel.Range)ws1.Cells[m_CellLoc[2] + iLoop2 + m_dtinfo.iH, 1];
                                        oRange2 = oRange2.EntireRow;
                                        oRange.Copy(oRange2);

                                        //oRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
                                    }
                                    else if (m_dtinfo.iV.Equals((int)1) ? true : false)
                                    {
                                        oRange = (Microsoft.Office.Interop.Excel.Range)ws1.Cells[1, m_CellLoc[3] + m_dtinfo.iV + iLoop2];
                                        oRange = oRange.EntireColumn;
                                        oRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                                    }
                                    break;
                                case "P":      // 새페이지.
                                    //출력 처리 혹은... sheet 복사 ? 
                                    iRepeatCnt = iRepeatCnt + 1;
                                    break;
                                case "R":      // 다른위치
                                    if ((iLoop2 + 1) % m_dtinfo.iMaxRow == 0)
                                        iRepeatCnt = iRepeatCnt + 1;
                                    break;
                                case "N":      // 없음. 해당데이터 종료.    
                                default:
                                    iRepeatCnt = iRepeatCnt + 1;
                                    break;
                            }//switch (m_dtinfo.sContinueMode)                       
                        }

                        if (iRepeatCnt >= m_dtinfo.iReCnt)
                            break;

                        for (int iLoop3 = 0; iLoop3 < dt.Columns.Count; iLoop3++) // 캡션
                        {
                            m_CellLoc = FindCell(ws1, dt.Columns[iLoop3].ColumnName, m_dtinfo.sOrientation, m_dtinfo.sHomeCell[iRepeatCnt]);

                            if (m_CellLoc[0] == 0) // 캡션을 못찾는경우 0으로 반환함. 건너뜀.
                                continue;

                            switch (dt.Columns[iLoop3].ColumnName.ToUpper())
                            {
                                case "QR":
                                case "QRCODE":
                                    //string tid = Convert.ToString(trSelectPrintContents.ManagedThreadId);
                                    //string startPath = Application.StartupPath + "\\";
                                    //QRCodeEncoder qr = new QRCodeEncoder();
                                    //qr.QRCodeEncodeMode = QRCodeEncoder.ENCODE_MODE.BYTE;
                                    //Image img = qr.Encode(dt.Rows[iLoop2][dt.Columns[iLoop3].ColumnName);
                                    //img.Save(startPath + tid + ".png");

                                    //ws1.Shapes.AddPicture(startPath + tid + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, fLeft, fTop, 150, 150);

                                    //File.Delete(startPath + tid + ".png");
                                    break;
                                case "이미지":
                                case "IMAGE":
                                    try
                                    {
                                        if (dt.Rows[iLoop2][dt.Columns[iLoop3].ColumnName] == null)
                                            break;

                                        string id = Convert.ToString(DateTime.Now.ToString("yyyyMMddhhmmss"));
                                        string startPath = Application.StartupPath + "\\";

                                        float fLeft = 0;
                                        float fTop = 0;
                                        float fHeight = 0;
                                        float fWidth = 0;

                                        oRange = (Microsoft.Office.Interop.Excel.Range)ws1.Cells[m_CellLoc[0], m_CellLoc[1]];
                                        oRange2 = (Microsoft.Office.Interop.Excel.Range)oRange.MergeArea;

                                        fLeft = (float)((double)oRange2.Left + (double)3);
                                        fTop = (float)((double)oRange2.Top + (double)3);
                                        fHeight = (float)((double)oRange2.Height - (double)6);
                                        fWidth = (float)((double)oRange2.Width - (double)6);

                                        byte[] bImage = (byte[])dt.Rows[iLoop2][dt.Columns[iLoop3].ColumnName];
                                        MemoryStream ms = new MemoryStream();
                                        Image img = null;
                                        ms.Position = 0;
                                        ms.Write(bImage, 0, (int)bImage.Length);
                                        img = Image.FromStream(ms);

                                        img.Save(startPath + id + ".png");
                                        ws1.Shapes.AddPicture(startPath + id + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, fLeft, fTop, fWidth, fHeight);
                                        File.Delete(startPath + id + ".png");
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception(ex.Message);
                                    }
                                    break;
                                default:
                                    // 할당. ( 캡션의 위치 + 방향값 + 증가값 - 최대값이후
                                    ws1.Cells[(m_dtinfo.iH.Equals((int)1) ? m_CellLoc[2] : m_CellLoc[0])
                                        + m_dtinfo.iH + (m_dtinfo.iH * iLoop2) - (m_dtinfo.iH * iRepeatCnt * m_dtinfo.iMaxRow)
                                                , (m_dtinfo.iV.Equals((int)1) ? m_CellLoc[3] : m_CellLoc[1])
                                                + m_dtinfo.iV + (m_dtinfo.iV * iLoop2) - (m_dtinfo.iV * iRepeatCnt * m_dtinfo.iMaxRow)]
                                                 = dt.Rows[iLoop2][dt.Columns[iLoop3].ColumnName];

                                    break;
                            } //    switch
                        } // for (int iLoop3 
                    }// for (int iLoop2
                }// for (int iLoop1              

                //-------------------------------
                //print 
                //excelApp.Visible = true;  //엑셀파일보기
                //excelApp.Sheets.PrintPreview(true); //미리보기모드 
                //ws1.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4; // A4로 설정
                //ws1.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape; //가로로 출력
                ws1.PrintOut(1, Type.Missing, 1, false, sPrintName);
                // (시작 페이지 번호, 마지막 페이지 번호, 출력 장 수, 프리뷰 활성 유/무, 활성프린터명, 파일로인쇄 (true), 여러장 한부씩 인쇄, 인쇄할 파일이름)                                              
                // PrintOut ?
                // object From /// 인쇄를 시작할 페이지 번호입니다. 이 인수를 생략하면 인쇄가 처음부터 시작됩니다.
                // , Object To     /// 인쇄할 마지막 페이지 번호입니다. 이 인수를 생략하면 마지막 페이지까지 인쇄됩니다.
                // , object Copies    /// 인쇄할 매수입니다. 이 인수를 생략하면 한 부만 인쇄됩니다.
                // , object Preview   /// Microsoft Office Excel에서 개체를 인쇄하기 전에 인쇄 미리 보기를 호출하려면 true이고, 개체를 즉시 인쇄하려면 false(또는 생략)입니다.
                // , object ActivePrinter  /// 활성 프린터의 이름을 설정합니다
                // , object PrintToFile  /// 파일로 인쇄하는 경우 true입니다. PrToFileName이 지정되지 않으면 Excel에서 출력 파일의 이름을 입력하라는 메시지를 표시합니다.
                // , object Collate   /// 여러 장을 한 부씩 인쇄하는 경우 true입니다.
                // , object PrToFileName  /// PrintToFile이 true로 설정되면 이 인수는 인쇄할 파일의 이름을 지정합니다.

                //엑셀 저장
                if (bSave)
                {
                    if (Directory.Exists(Application.StartupPath + @"\EXCEL\") == false)
                        Directory.CreateDirectory(Application.StartupPath + @"\EXCEL\");
                    string sFilePath = Application.StartupPath + @"\EXCEL\" + Path.GetFileName(this.sFileName).Split('.')[0] + DateTime.Now.ToString("_yyyyMMddHHmmss");
                    ws1.SaveAs(sFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                uint processId = 0;
                GetWindowThreadProcessId(new IntPtr(excelApp.Hwnd), out processId);

                if (wb != null)
                {
                    // Clean up
                    wb.Close(0);
                    ReleaseExcelObject(ws1);
                    ReleaseExcelObject(wb);
                }

                excelApp.Quit();
                ReleaseExcelObject(excelApp);
                excelApp = null;
                wb = null;

                if (processId != 0)
                {
                    System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)processId);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }
            }
        }//PrintReport

    }//class
}//namespace