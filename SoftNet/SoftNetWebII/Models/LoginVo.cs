using Base.Services;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;

namespace SoftNetWebII.Models
{
    public class StoreList
    {
        public string S_MFNO { get; set; } = "";
        public string S_Station { get; set; } = "";
        public string S_PartNO { get; set; } = "";
        public string ActionType { get; set; } = "";
        public string Station { get; set; } = "";
        public string StoreNO { get; set; } = "";
        public string StoreSpacesNO { get; set; } = "";
        public string Select_ID { get; set; } = "";
        public string Select_ID_QTY { get; set; } = "";
        public string Select_DroneType { get; set; } = "1";
        public string ERRMsg { get; set; } = "";
        
        public DataTable SIDList { get; set; }
        public DataTable New_K_List { get; set; }
        public DataTable OLD_K_List { get; set; }
		public DataTable OLD_K_IN_Station_List { get; set; }
		public List<string[]> Station_List { get; set; }
        public List<string> StoreSpacesNO_List { get; set; }
        public DataTable TotalStock_List { get; set; }
        public string DisplayHTML { get; set; } = "";
    }
    public class LabelDOC3stockII
    {
        public LabelDOC3stockII() { }
        public LabelDOC3stockII (char Id_Type,string Id, string DOCNumberNO, string PartNO, string Price, string Unit, string QTY, string Remark, string SimulationId, string IsOK, string IN_StoreNO, string IN_StoreSpacesNO, string OUT_StoreNO, string OUT_StoreSpacesNO, string ArrivalDate)
        {
            this.Id_Type = Id_Type;
            this.Check = true;
            this.Id= Id;
            this.DOCNumberNO = DOCNumberNO;
            this.PartNO= PartNO;
            this.Price= Price;
            this.Unit= Unit;
            this.QTY= QTY;
            this.Remark= Remark;
            this.SimulationId = SimulationId;
            this.IsOK= IsOK;
            this.IN_StoreNO= IN_StoreNO;
            this.IN_StoreSpacesNO= IN_StoreSpacesNO;
            this.OUT_StoreNO= OUT_StoreNO;
            this.OUT_StoreSpacesNO = OUT_StoreSpacesNO;
            this.ArrivalDate= ArrivalDate;
        }
        public char Id_Type { get; set; }
        public bool Check { get; set; }
        public string Id { get; set; }
        public string DOCNumberNO { get; set; }
        public string PartNO { get; set; }
        public string Price { get; set; }
        public string Unit { get; set; }
        public string QTY { get; set; }
        public string Remark { get; set; }
        public string SimulationId { get; set; }
        public string IsOK { get; set; }
        public string IN_StoreNO { get; set; }
        public string IN_StoreSpacesNO { get; set; }
        public string OUT_StoreNO { get; set; }
        public string OUT_StoreSpacesNO { get; set; }
        public string ArrivalDate { get; set; }
    }

/*
    public class LabelWork
    {
        public string Station { get; set; } = "";
        public string Station_Name { get; set; } = "";
        public string StationNO_Custom_DisplayName { get; set; } = "";
        public string Type { get; set; }
        public string TypeValue { get; set; }
        public string OrderNO { get; set; } = "";
        public string SimulationId { get; set; } = "";
        public string IndexSN { get; set; } = "";
        public string OPNO { get; set; }
        public string OPNO_Name { get; set; }
        public int OKQTY { get; set; }
        public int FailQTY { get; set; }
        public string State { get; set; } = "";
        public string Has_Knives { get; set; } = "0";
        public string LocalIPPort { get; set; } = "";
        public string ERRMsg { get; set; } = "";
        public List<string[]> MES_List { get; set; }
        public string MES_String { get; set; } = "";
        public string Station_State { get; set; } = "";
        public string RemarkTimeS { get; set; } = "";
        public string RemarkTimeE { get; set; } = "";
        public string WO_QTY { get; set; } = "0";
        public string TOT_OK_QTY { get; set; } = "0";
        public string TOT_Fail_QTY { get; set; } = "0";
        public string Transfer_QTY { get; set; } = "0";
        public string PartNO_INFO { get; set; } = "";
        public string WO_CT { get; set; } = "0";
        public string E_CT { get; set; } = "0";
        public string Station_Config_Store_Type { get; set; } = "1";
        public string BOOLonToggleMenu { get; set; } = "";

        public List<string[]> HasWO_List { get; set; }
        public List<string[]> HasPO_List { get; set; }

        

    }
 

    public class OddCT_GetData
    {
        public string Key { get; set; } = "";
        public string Value { get; set; } = "";
    }
    public class SystemConfigObj
    {
        //監看Thead狀態  0=SfcTimerloopUpdateTagValue_Tick 1=SfcTimerloopautoRUN_DOC_Tick 2=SfcTimerloopautoRUN_Json_Tick 3=SfcTimerloopthread_Tick 4=DeviceConnectCheck_Tick

        public  long _a01 { get; } = _Fun._a01;
        public  long _a02 { get; } = _Fun._a02;
        public long _a03 { get; } = _Fun._a03;
        public long _a04 { get; } = _Fun._a04;
        public long _a05 { get; } = _Fun._a05;
        public long _a06 { get; } = _Fun._a06;
        public  long _a07 { get; } = _Fun._a07;
        public long _a08 { get; } = _Fun._a08;
        public long _a09 { get; } = _Fun._a09;
        public long _a10 { get; } = _Fun._a10;
        public  long _a11 { get; } = _Fun._a11;
        public  long _a12 { get; } = _Fun._a12;
        public  long _a13 { get; } = _Fun._a13;
        public  long _a14 { get; } = _Fun._a14;
        public long _a15 { get; } = _Fun._a15;
        public  long _a16 { get; } = _Fun._a16;

        public bool T0 { get; } = _Fun.Is_RUNTimeServer_Thread_State[0];
        public bool T1 { get;  } = _Fun.Is_RUNTimeServer_Thread_State[1];
        public bool T2 { get; } = _Fun.Is_RUNTimeServer_Thread_State[2];
        public bool T3 { get; } = _Fun.Is_RUNTimeServer_Thread_State[3];
        public bool T4 { get; } = _Fun.Is_RUNTimeServer_Thread_State[4];
        public bool Thread_ForceClose { get; } = _Fun.Is_Thread_ForceClose;
        public bool Is_Thread_For_Test { get; } = _Fun.Is_Thread_For_Test;

        public string Send_RUNTimeServer_Thread_State { get; set; } = "";

        public string FunType { get; set; } = "";
    }
    public class LabelProject
    {
        public string PP_Name { get; set; }
        public string PartNO { get; set; } = "";
        public string PartName { get; set; } = "";
        public string Specification { get; set; } = "";
        public string StationNO { get; set; } = "";
        public string StationName { get; set; }
        public string Type { get; set; }
        public string OrderNO { get; set; }
        public string IndexSN { get; set; } = "0";
        public string OPNO { get; set; } = "";
        public string OPNO_Name { get; set; }
        public string State { get; set; } = "";

        public int TOTALOKQTY { get; set; } = 0;
        public int TOTALFailQTY { get; set; } = 0;

        public string TOTALWorkTime { get; set; } = "";
        public int OKQTY { get; set; } = 0;
        public int FailQTY { get; set; } = 0;
        public string LocalIPPort { get; set; } = "";
        public string ERRMsg { get; set; } = "";
        public List<string[]> MES_List { get; set; }
        public string MES_String { get; set; } = "";
        public string Station_State { get; set; } = "";

        public List<string[]> HasPP_Name_List { get; set; }
        public List<string[]> HasPO_List { get; set; }
        public List<string[]> HasPartNO_List { get; set; }
        public string Remark { get; set; } = "";

    }
    /*
    public class MUTIStationObj
    {
        public string OrderBY_CMD { get; set; } = "order by PartNO,IndexSN,OrderNO,RemarkTimeS desc";
        public string SI_PP_Name { get; set; } = "";
        public string SI_PartNO { get; set; } = "";
        public string SI_PartName { get; set; } = "";
        public string SI_OrderNO { get; set; } = "";
        public string SI_IndexSN { get; set; } = "";
        public int SI_OKQTY { get; set; }
        public int SI_FailQTY { get; set; }
        public string SI_OPNO { get; set; } = "";
        public string SI_Slect_OPNOs { get; set; } = "";
        public string SI_SimulationId { get; set; } = "";
        public string Has_Knives { get; set; } = "0";
        public string State { get; set; } = "";
        public string StationNO { get; set; } = "";
        public string StationName { get; set; }
        public string LocalIPPort { get; set; } = "";
        public List<string[]> HasPP_Name_List { get; set; }
        public List<string[]> HasPartNO_List { get; set; }
        public List<string[]> HasWorkData_List { get; set; }
        public List<string[]> HasPO_List { get; set; }
        public List<string[]> HasWO_List { get; set; }
        public string Station_Config_Store_Type { get; set; } = "1";

        public string Remark { get; set; } = "";
        public string MES_String { get; set; } = "";
        public string MES_Report { get; set; } = "";
        public string ERRMsg { get; set; } = "";
        public string Select_ID { get; set; } = "";
        public List<string[]> StationNOList { get; set; }
        public bool OutError { get; set; } = false;

    }
    */
    public class APSViewBOM
    {
        public string Fun_S_List { get; set; } = "";
        public string NewPartNO { get; set; } = "";
        public string MBOMId { get; set; } = "";//母bom ID
        public string BOMId { get; set; } = "";//本階bom ID
        public string SelectBOMId { get; set; } = "";//畫面選擇的母bom
        public string SelectPP_DATA { get; set; } = "";//異動的製程
        public string SimulationId { get; set; } = "";
        public string ERRMsg { get; set; } = "";

    }

    public class Z_CreateBOM_Data
    {
        public string IPPort { get; set; } = "";
        public string StationNO { get; set; } = "";
        public int QTY { get; set; } = 0;
        public float Weight { get; set; } = 0f;
        public string MES_Report { get; set; } = "";
        public string ERRMsg { get; set; } = "";
        public List<string[]> StationNOList { get; set; }

    }
    public class CreateBOMOBJ
    {
        public string COMType { get; set; } = "";
        public string SClass { get; set; } = "";
        public string SPartNO { get; set; } = "";
        public string SPartName { get; set; } = "";
        public string SApply_PartNO { get; set; } = "";
        public string SPP_Name { get; set; } = "";
        public string SCOPY_PP_Name { get; set; } = "";
        public string SApply_StationNO { get; set; } = "";
        public string SStationNO_Custom_DisplayName { get; set; } = "";
        public string SStation_DIS_Remark { get; set; } = "";
        public int IndexSN { get; set; } = 0;
        public int SBOMQTY { get; set; } = 1;
        public string SStationNO_IndexSN_Merge { get; set; } = "";
        public int SIsChackQTY { get; set; } = 0;
        public int SIsChackIsOK { get; set; } = 0;
        public string MFNO { get; set; } = "";
        public string Fun_S_List { get; set; } = "";
        public string MBOMId { get; set; } = "";//母bom ID
        public string BOMId { get; set; } = "";//本階bom ID
        public string SelectBOMId { get; set; } = "";//畫面選擇的母bom
        public string SelectPP_DATA { get; set; } = "";//異動的製程
        public string ERRMsg { get; set; } = "";
        public List<string> NeedIdList { get; set; }
        public List<string[]> HasClass_List { get; set; }
        public List<string> HasPartNO_List { get; set; }
        public List<string[]> HasStationNO_List { get; set; }
        public List<string[]> HasMFNO_List { get; set; }

    }
    public class APSViewData
    {
        public string Fun_S_List { get; set; } = "";
        public string NewPartNO { get; set; } = "";
        public string MBOMId { get; set; } = "";
        public string SelectNeedId { get; set; } = "";
        public string SelectFun1 { get; set; } = "";
        public string SelectFun2 { get; set; } = "A3";
        public string SelectFun3 { get; set; } = "B1";
        public string SelectFun4 { get; set; } = "";
        public string SelectFun5 { get; set; } = "";
        public string SelectFun6 { get; set; } = "";
        public string ERRMsg { get; set; } = "";
        public List<string[]> NeedIdList { get; set; }
    }
    public class LabelCreateBarCode
    {
        public int Width { get; set; } = 300;
        public int Height { get; set; } = 300;
        public string StationNO { get; set; }
        public List<string[]> StationNOList { get; set; }
    }

        public class LoginVo
    {
        public string Account { get; set; }
        public string Pwd { get; set; }
        //public string Locale { get; set; }

        public string FromUrl { get; set; }

        public string AccountMsg { get; set; }
        public string PwdMsg { get; set; }
        public string ErrorMsg { get; set; }
    }

    public class SignRowDto
    {
        public string NodeName { get; set; }
        public string SignerName { get; set; }
        //public string Locale { get; set; }

        public string SignStatusName { get; set; }

        public string SignTime { get; set; }
        public string Note { get; set; }
    }
}