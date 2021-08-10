#region < HEADER AREA >
/*--------------------------------------------------------------------------------------------
CREATE       : 2021-06-10 JEJUN  
DESCRIPT    :  엑셀리포트헬퍼 - 리포트 형식별 테이블정보 조회하여 레포트처리클래스로 넘김.

UPDATE       :
DESCRIPT    :  
---------------------------------------------------------------------------------------------*/
#endregion

#region < USING AREA >
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

using System.Windows.Forms;
#endregion

namespace ExcelReportHelper
{
    public class ExcelReportHelper
    {
        #region < 멤버 변수 >
        private string sKeyword;
        private DataSet ds;
        private string sInfoWhere;
        private string sDataWhere;

        private string sFileName;
        private string sProcedure;

        private List<ExeclReportDtInfo> dtinfo;
        #endregion

        #region < 생성자 >
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sKeyword">  출력 키워드 </param>
        /// <param name="sKey"> 레포트유형 중 각 조건별 출력 파일네임 다를경우, 조건값.  </param>
        /// /// <param name="drWhere"> 데이터 조회값  </param>
        public ExcelReportHelper(string sKeyword, DataRow drInfoWhere, DataRow drDataWhere)
        {
            this.sKeyword = sKeyword;
            this.sInfoWhere = MakeParmString(drInfoWhere);
            this.sDataWhere = MakeParmString(drDataWhere);
            //초기화
            dtinfo = new List<ExeclReportDtInfo>();
            GetReportInfo();
        }
        public ExcelReportHelper(string sKeyword, DataRow drWhere)
        {
            this.sKeyword = sKeyword;
            this.sDataWhere = MakeParmString(drWhere);
            //초기화
            dtinfo = new List<ExeclReportDtInfo>();
            GetReportInfo();
        }
        #endregion


        /// <summary>
        /// datarow => string
        /// 규칙 : column1 = 값1; column2 = 값2;  ...
        /// </summary>
        /// <param name="dr"></param>
        private string MakeParmString(DataRow dr)
        {
            try
            {
                // dataraw => string
                string sParameter = string.Empty;

                for (int iloop = 0; iloop < dr.ItemArray.Length; iloop++)
                    sParameter = string.Concat(sParameter, dr.Table.Columns[iloop].ColumnName, "=", dr.ItemArray.GetValue(iloop), ";");

                return sParameter;
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 레포트 유형에서도 각 조건별 출력해야할 파일이름과 데이터그룹정보가져옴. 
        /// </summary>
        private void GetReportInfo()
        {
            // 레포트기준정보 조회
            DBHelper helper;
            helper = new DBHelper(false);
            try
            {
                DataSet dstemp = helper.FillDataSet("EXCELREPORTINFO_S", CommandType.StoredProcedure
                    , helper.CreateParameter("AS_PLANTCODE", DBHelper.nvlString(LoginInfo.PlantCode), DbType.String, ParameterDirection.Input)
                       , helper.CreateParameter("AS_KEYWORD", DBHelper.nvlString(sKeyword), DbType.String, ParameterDirection.Input)
                         //, helper.CreateParameter("AS_PARAMETERS", DBHelper.nvlString(sKey), DbType.String, ParameterDirection.Input)
                         , helper.CreateParameter("RS_FILENAME", DbType.String, ParameterDirection.Output, null, 100)
                         , helper.CreateParameter("RS_PROCNAME", DbType.String, ParameterDirection.Output, null, 100)
                       );

                if (helper.RSCODE == "S")
                {
                    if (dstemp.Tables.Count <= 0)
                    {
                        throw new Exception("EXCELREPORTINFO_S " + " NO DATA!!");
                    }

                    this.sFileName = DBHelper.nvlString(helper.Parameters["RS_FILENAME"].Value);
                    this.sProcedure = DBHelper.nvlString(helper.Parameters["RS_PROCNAME"].Value);

                    for (int iloop = 0; iloop < dstemp.Tables[0].Rows.Count; iloop++)
                    {
                        ExeclReportDtInfo dtInfotemp = new ExeclReportDtInfo(DBHelper.nvlInt(dstemp.Tables[0].Rows[iloop]["TABLEGROUP"])
                                                                                                         , DBHelper.nvlString(dstemp.Tables[0].Rows[iloop]["ORIENTATION"])
                                                                                                        , DBHelper.nvlString(dstemp.Tables[0].Rows[iloop]["HOMECELL"])
                                                                                                        , DBHelper.nvlInt(dstemp.Tables[0].Rows[iloop]["MAXIMUM"])
                                                                                                        , DBHelper.nvlString(dstemp.Tables[0].Rows[iloop]["CONTINUEMODE"])
                                                                                                        , DBHelper.nvlInt(dstemp.Tables[0].Rows[iloop]["REPEATCOUNT"])
                                                                                                        );
                        dtinfo.Add(dtInfotemp);
                    }

                    // 데이터셋 조회
                    ds = helper.FillDataSet(this.sProcedure, CommandType.StoredProcedure
                   , helper.CreateParameter("AS_PARAMETER", this.sDataWhere, DbType.String, ParameterDirection.Input)
                   );

                    if (helper.RSCODE == "S")
                    {
                        if (ds.Tables.Count <= 0)
                        {
                            throw new Exception(this.sProcedure + " NO DATA!!");
                        }
                    }
                    else
                    {
                        throw new Exception(helper.RSMSG);
                    }
                }
                else
                {
                    throw new Exception(helper.RSMSG);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// 헬퍼에서 프린트 진행 시 , 레포트처리 호출
        /// </summary>
        public void Print()
        {
            try
            {
                string sFilePath = string.Empty;
                sFilePath = Application.StartupPath + "\\EXCEL\\" + this.sFileName;
                FileInfo fileInfo = new FileInfo(sFilePath);
                if (!fileInfo.Exists)
                {
                    sFilePath = Application.StartupPath + "\\" + this.sFileName;
                    fileInfo = new FileInfo(sFilePath);
                    if (!fileInfo.Exists)
                    {
                        throw new Exception(sFilePath + " 파일을 찾을 수 없습니다");
                    }
                }

                string sPrintName = string.Empty;
                PrintDialog printDialog = new PrintDialog();
                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    sPrintName = printDialog.PrinterSettings.PrinterName;
                    WIZ.REPORT.ExcelReport excelreport = new REPORT.ExcelReport(sFilePath, this.dtinfo, this.ds);
                    excelreport.PrintReport(sPrintName);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

    }//class
}//namespace