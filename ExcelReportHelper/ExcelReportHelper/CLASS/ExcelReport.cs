﻿#region < HEADER AREA >
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
using System.Text.RegularExpressions;
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
        object[,] rngData;

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
        internal ExcelReport(string sFileName, List<ExeclReportDtInfo> execlReportDtInfo,  DataSet dataSet)
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
            int[] returnobj = new int[] { 1, 1 };
            // x 숫자짜르기
            string sX = Regex.Replace(s, @"\D", "");//숫자 추출
            returnobj[0] = string.IsNullOrEmpty(sX) ? 1 : Convert.ToInt32(sX);
            // y 
            string sY = Regex.Replace(s, @"\d", "");//문자 추출
            returnobj[1] = iTransAlpha(sY);

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
        static int[] FindCell(object[,] array, string elem, string orientation, string home)
        {
            int[] returnVal = new int[2] { 0, 0 };

            int rowEndIndx = array.GetLength(0);
            int colEndIndx = array.GetLength(1);

            int[] homeindex = CellToXY(home);
             int rowStartIndx = homeindex[0];
            int colStartIndx = homeindex[1];

            if (orientation.Equals( "V"))
                colEndIndx = colStartIndx;
            if (orientation.Equals("H"))
                rowEndIndx = rowStartIndx;
            else if (orientation.Equals("N"))
                return homeindex;

            for (int rowIndex = rowStartIndx; rowIndex <= rowEndIndx; rowIndex++)
            {
                for (int colIndex = colStartIndx; colIndex <= colEndIndx; colIndex++)
                {
                    if (Convert.ToString(array[rowIndex, colIndex]) == elem)
                    {
                        returnVal.SetValue(rowIndex, 0);
                        returnVal.SetValue(colIndex, 1);
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
        internal void PrintReport(string sPrintName)
        {
            try
            {
                int iRepeatCnt = 0;  // 테이블 반복 횟수 
                //-------------------------------
                // 엑셀 오픈                   
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                wb = excelApp.Workbooks.Open(this.sFileName);
                ws1 = wb.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet;
                Microsoft.Office.Interop.Excel.Range rng = ws1.UsedRange;   //현재 시트에서 사용중인 범위
                rngData = rng.Value; //범위의 데이터

                //-------------------------------
                for (int iLoop1 = 0; iLoop1 < dtinfo.Count ; iLoop1++) // 데이터테이블
                {
                    iRepeatCnt = 0;
                    dt = ds.Tables[iLoop1];
                    m_dtinfo = dtinfo[iLoop1];
                    for (int iLoop2 = 0; iLoop2 < dt.Rows.Count; iLoop2++) // 로우
                    {
                        if (iRepeatCnt >= m_dtinfo.iReCnt) 
                            break;

                        for (int iLoop3 = 0; iLoop3 < dt.Columns.Count; iLoop3++) // 캡션
                        {
                            m_CellLoc = FindCell(rngData, dt.Columns[iLoop3].ColumnName, m_dtinfo.sOrientation, m_dtinfo.sHomeCell[iRepeatCnt]);
                            if (m_CellLoc[0] == 0) // 캡션을 못찾는경우 0으로 반환함. 건너뜀.
                                continue;
                            // 할당. ( 캡션의 위치 + 방향값 + 줄바꿈반영값 
                            ws1.Cells[m_CellLoc[0] + m_dtinfo.iH + (m_dtinfo.iH * (iLoop2 % m_dtinfo.iMaxRow))
                                            , m_CellLoc[1] + m_dtinfo.iV + (m_dtinfo.iV * (iLoop2 % m_dtinfo.iMaxRow))]
                                             = dt.Rows[iLoop2][dt.Columns[iLoop3].ColumnName];
                        }

                        //max행 넘으면 줄바꿈
                        if ((iLoop2 + 1)  % m_dtinfo.iMaxRow == 0) 
                            iRepeatCnt = iRepeatCnt + 1;
                    }
                }

                //-------------------------------
                //print 
                //excelApp.Visible = true;  //엑셀파일보기
                //excelApp.Sheets.PrintPreview(true); //미리보기모드 
                //ws1.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4; // A4로 설정
                //ws1.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape; //가로로 출력

                ws1.PrintOut(1, 1, 1, false, sPrintName); 
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