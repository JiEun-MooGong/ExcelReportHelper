#region < HEADER AREA >
/*--------------------------------------------------------------------------------------------
CREATE       : 2021-06-10 JEJUN  
DESCRIPT    : 레포트데이터그룹정보 클래스 
---------------------------------------------------------------------------------------------*/
#endregion

#region < USING AREA >
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#endregion

namespace ExcelReportHelper
{
    internal class ExeclReportDtInfo         //데이터 그룹 정보
    {
        private int _iId;                        // TABLEGROUP	
        private int _iH;                         //1 : ORIENTATION = H       => 방향별 증가 처리를 위한 값 
        private int _iV;                         //1 : ORIENTATION = V / 0 : ORIENTATION : H 
        private string _sOrientation;  //1 : ORIENTATION = V  H 
        private string[] _sHomeCell;      //HOMECELL 
        private int _iMaxRow;            //MAXIMUM	
        private string _sContinueMode;        //CONTINUEMODE
        private int _iReCnt;            //REPEATCOUNT		

        internal ExeclReportDtInfo(int iId, string sOrientation, string sHomeCell, int iMaxRow, string sContinueMode, int iReCnt)
        {
            this._iId = iId;
            this._sOrientation = sOrientation;
            switch (sOrientation)
            {
                case "V":
                    this._iH = 0;
                    this._iV = 1;
                    break;
                case "H":
                    this._iH = 1;
                    this._iV = 0;
                    break;
                default:
                    this._iH = 0;
                    this._iV = 0;
                    break;
            }
            this._sHomeCell = sHomeCell.Split(';');
            this._iMaxRow = iMaxRow;
            this._sContinueMode = sContinueMode;
            this._iReCnt = iReCnt;
        }

        internal int iId
        {
            get { return this._iId; }
        }
        internal int iH
        {
            get { return this._iH; }
        }
        internal int iV
        {
            get { return this._iV; }
        }
        internal string sOrientation
        {
            get { return this._sOrientation; }
        }
        internal string[] sHomeCell
        {
            get { return this._sHomeCell; }
        }
        internal int iMaxRow
        {
            get { return this._iMaxRow; }
        }
        internal string sContinueMode
        {
            get { return this._sContinueMode; }
        }
        internal int iReCnt
        {
            get { return this._iReCnt; }
        }

    }//class
}//namespace
