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
        DBHelper helper;

        private string sKeyword;
        private DataSet dsData;
        private string sInfoWhere;
        private string sDataWhere;

        private string sPrintName;
        private string sFileName;
        private string sFilePath;
        private bool bSave;
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
        public ExcelReportHelper(string sKeyword)
        {
            try
            {
                this.sKeyword = sKeyword;
                GetReportInfo();

                this.sFilePath = Application.StartupPath + "\\EXCEL\\" + this.sFileName;
                FileInfo fileInfo = new FileInfo(this.sFilePath);
                if (!fileInfo.Exists)
                {
                    this.sFilePath = Application.StartupPath + "\\" + this.sFileName;
                    fileInfo = new FileInfo(this.sFilePath);
                    if (!fileInfo.Exists)
                        throw new Exception(this.sFilePath + " 파일을 찾을 수 없습니다");
                }

                PrintDialog printDialog = new PrintDialog();
                if (printDialog.ShowDialog() == DialogResult.OK)
                {
                    this.sPrintName = printDialog.PrinterSettings.PrinterName;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
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

        private void GetReportInfo()
        {
            // 레포트기준정보 조회
            helper = new DBHelper(false);
            this.dtinfo = new List<ExeclReportDtInfo>();
            try
            {
                DataSet dstemp = helper.FillDataSet("EXCELREPORTINFO_S", CommandType.StoredProcedure
                    , helper.CreateParameter("AS_PLANTCODE", DBHelper.nvlString(LoginInfo.PlantCode), DbType.String, ParameterDirection.Input)
                       , helper.CreateParameter("AS_KEYWORD", DBHelper.nvlString(this.sKeyword), DbType.String, ParameterDirection.Input)
                         //, helper.CreateParameter("AS_PARAMETERS", DBHelper.nvlString(sKey), DbType.String, ParameterDirection.Input)
                         , helper.CreateParameter("RS_FILENAME", DbType.String, ParameterDirection.Output, null, 100)
                         , helper.CreateParameter("RS_PROCNAME", DbType.String, ParameterDirection.Output, null, 100)
                         , helper.CreateParameter("RS_SAVEFLAG", DbType.String, ParameterDirection.Output, null, 100)
                       );

                if (helper.RSCODE != "S")
                    throw new Exception(helper.RSMSG);

                if (dstemp.Tables.Count <= 0)
                    throw new Exception("EXCELREPORTINFO_S " + " NO DATA!!");

                this.sFileName = DBHelper.nvlString(helper.Parameters["RS_FILENAME"].Value);
                this.sProcedure = DBHelper.nvlString(helper.Parameters["RS_PROCNAME"].Value);
                this.bSave = DBHelper.nvlString(helper.Parameters["RS_SAVEFLAG"].Value).Equals("Y") ? true : false;

                for (int iloop = 0; iloop < dstemp.Tables[0].Rows.Count; iloop++)
                {
                    ExeclReportDtInfo dtInfotemp = new ExeclReportDtInfo(DBHelper.nvlInt(dstemp.Tables[0].Rows[iloop]["TABLEGROUP"])
                                                                                                     , DBHelper.nvlString(dstemp.Tables[0].Rows[iloop]["ORIENTATION"])
                                                                                                    , DBHelper.nvlString(dstemp.Tables[0].Rows[iloop]["HOMECELL"])
                                                                                                    , DBHelper.nvlInt(dstemp.Tables[0].Rows[iloop]["MAXIMUM"])
                                                                                                    , DBHelper.nvlString(dstemp.Tables[0].Rows[iloop]["CONTINUEMODE"])
                                                                                                    , DBHelper.nvlInt(dstemp.Tables[0].Rows[iloop]["REPEATCOUNT"])
                                                                                                    );
                    this.dtinfo.Add(dtInfotemp);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void GetData(DataRow drWhere)
        {
            try
            {
                this.sDataWhere = string.Empty;
                this.sDataWhere = MakeParmString(drWhere);
                this.GetData();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void GetData()
        {
            helper = new DBHelper(false);
            try
            {
                // 데이터셋 조회
                dsData = helper.FillDataSet(this.sProcedure, CommandType.StoredProcedure
               , helper.CreateParameter("AS_PARAMETER", this.sDataWhere, DbType.String, ParameterDirection.Input)
               );

                if (helper.RSCODE == "S")
                {
                    if (dsData.Tables.Count <= 0)
                        throw new Exception(this.sProcedure + " NO DATA!!");
                }
                else
                    throw new Exception(helper.RSMSG);
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
            WIZ.REPORT.ExcelReport excelreport = new REPORT.ExcelReport(this.sFilePath, this.dtinfo, this.dsData);
            excelreport.PrintReport(this.sPrintName, this.bSave);
        }

    }//class
}//namespace