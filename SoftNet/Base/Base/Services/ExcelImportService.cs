using Base.Models;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

//excel import
namespace Base.Services
{
    public class ExcelImportService<T> where T : class, new()
    {
        //constant
        const string RowSep = "\r\n";  //row seperator
        
        //ok excel row no
        private List<int> _okRowNos = new List<int>();

        //failed excel row no/msg
        private List<SnStrDto> _failRows = new List<SnStrDto>();

        //cell x-way name(no number)
        private string CellXname(string colName)
        {
            return Regex.Replace(colName, @"[\d]", string.Empty);
        }

        /// <summary>
        /// import by stream
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="importDto"></param>
        /// <param name="fileName"></param>
        /// <param name="uiDtFormat"></param>
        /// <returns></returns>
        public async Task<ResultImportDto> ImportByStreamAsync(Stream stream, ExcelImportDto<T> importDto, string dirUpload, string fileName, string uiDtFormat)
        {
            stream.Position = 0;
            var docx = _Excel.StreamToDocx(stream);
            var result = await ImportByDocxAsync(docx, importDto, dirUpload, fileName, uiDtFormat);

            //release docx
            docx = null;
            return result;
        }



        /// <summary>
        /// import by excel docx
        /// excel row/cell (base 0)
        /// </summary>
        /// <param name="docx"></param>
        /// <param name="importDto"></param>
        /// <param name="fileName">imported excel file name</param>
        /// <param name="uiDtFormat"></param>
        /// <returns>error msg if any</returns>
        public async Task<ResultImportDto> ImportByDocxAsync(SpreadsheetDocument docx, ExcelImportDto<T> importDto, string dirUpload, string fileName, string uiDtFormat)
        {
            uiDtFormat = "yyyy-MM-dd HH:mm:ss";
            ResultImportDto re_DATA = new ResultImportDto();

            #region 1.set variables
            #region set docx, excelRows, ssTable
            //var errorMsg = "";
                var wbPart = docx.WorkbookPart;
            var wsPart = (WorksheetPart)wbPart.GetPartById(
                wbPart.Workbook.Descendants<Sheet>().ElementAt(importDto.SheetNo).Id);

            var excelRows = wsPart.Worksheet.Descendants<Row>();    //include empty rows
            var ssTable = wbPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable;
            #endregion

            #region set importDto.ExcelFids, excelFidLen
            int idx;
            var colMap = new JObject();     //col x-way name(ex:A) -> col index
            var cells = excelRows.ElementAt(importDto.FidRowNo - 1).Elements<Cell>();
            var excelFids = new List<string>(); //欄位名稱
            DataTable XMLElementDataTable = new DataTable();
            //if (importDto.ExcelFids == null || importDto.ExcelFids.Count == 0)
            //{
            //如果沒有傳入excel欄位名稱, 則使用第一行excel做為欄位名稱
                idx = 0;
                foreach (var cell in cells)
                {
                    excelFids.Add(GetCellValue(ssTable, cell));
                    colMap[CellXname(cell.CellReference)] = idx;
                    idx++;
                }

            bool is_ckeck_cell = true;
            foreach (var cell in excelFids)
            {
                XMLElementDataTable.Columns.Add(new DataColumn(cell, typeof(string)));
                if (is_ckeck_cell)
                {
                    switch (importDto.ImportType)
                    {
                        case "Simulation"://排程  
                            if (cell != "是否要轉") { re_DATA.ErrorMsg = "Excel檔, 格式不正確!.";return re_DATA; }
                            break;
                        case "Material":
                            break;
                        case "PPName_AND_BOM":
                            break;
                    }
                    is_ckeck_cell = false;
                }
            }

            int fno;
            var modelFids = new List<string>();         //全部欄位
            var model = new T();
            foreach (var prop in model.GetType().GetProperties())
            {
                //如果對應的excel欄位不存在, 則不記錄此欄位(skip)
                //var type = prop.GetValue(model, null).GetType();
                var fid = prop.Name;
                fno = excelFids.FindIndex(a => a == fid);
                if (fno < 0)
                    continue;

                modelFids.Add(fid);
                //excelIsDates[fno] = true;
            }



            for (var i = importDto.FidRowNo; i < excelRows.LongCount(); i++)
            {
                var excelRow = excelRows.ElementAt(i);
                var fileRow = new T();
                DataRow _dr2 = XMLElementDataTable.NewRow();
                int col = -1;
                foreach (Cell cell in excelRow)
                {

                    try
                    {
                        var value = "";
                        if (cell.CellReference == null) 
                        {
                            continue;
                        }
                        fno = (int)colMap[CellXname(cell.CellReference)];
                        if (cell.DataType != null)
                        {
                            value = (cell.DataType == CellValues.SharedString)
                                ? ssTable.ChildElements[int.Parse(cell.CellValue.Text)].InnerText
                                : cell.CellValue.Text;
                            _dr2[fno] = value;
                        }
                        else 
                        {
                            if (cell.CellValue != null)
                            { _dr2[fno] = cell.CellValue.Text; }
                            else
                            { _dr2[fno] = ""; }
                        }
                        if (_dr2[fno].ToString().Trim() != "")
                        {
                            switch (importDto.ImportType)
                            {
                                case "Simulation"://排程   0=是否要轉  1=訂單編號 2=客戶編號 3=客戶名稱 4=是否為補足安全量 5=適用行事曆編號 6=預計需求日 7=需求日容許正負差(小時) 8=物料編號 9=需求量	10=適用BOM編號 11=適用生產製程 12=適用工廠編號
                                    if (fno == 6) { _dr2[fno] = DateTime.FromOADate(double.Parse(_dr2[fno].ToString())).ToString(uiDtFormat); }
                                    break;
                                case "Material":
                                    break;
                                case "PPName_AND_BOM":
                                    break;
                            }
                        }
                        /*
                        if (col == 2)
                        {
                            if (cell.DataType != null)
                            {
                                var value = (cell.DataType == CellValues.SharedString)
                                    ? ssTable.ChildElements[int.Parse(cell.CellValue.Text)].InnerText
                                    : cell.CellValue.Text;
                                _dr2[++col] = value;
                            }
                            else { _dr2[++col] = cell.CellValue.Text; }
                        }
                        else if (col == 5)
                        {
                            if (cell.DataType != null)
                            {
                                var value = (cell.DataType == CellValues.SharedString)
                                    ? ssTable.ChildElements[int.Parse(cell.CellValue.Text)].InnerText
                                    : cell.CellValue.Text;
                                _dr2[++col] = DateTime.FromOADate(double.Parse(value)).ToString(uiDtFormat);
                            }
                            else { _dr2[++col] = DateTime.FromOADate(double.Parse(cell.CellValue.Text)).ToString(uiDtFormat); }
                        }
                        else
                        {
                            if (cell.DataType != null)
                            {
                                var value = (cell.DataType == CellValues.SharedString)
                                    ? ssTable.ChildElements[int.Parse(cell.CellValue.Text)].InnerText
                                    : cell.CellValue.Text;
                                _dr2[++col] = value;
                            }
                            else { _dr2[++col] = cell.CellValue.Text; }
                        }
                        */
                    }
                    catch (Exception ex)
                    {
                        string _s = "";
                    }
                }
                XMLElementDataTable.Rows.Add(_dr2);
            }


            if (XMLElementDataTable != null && XMLElementDataTable.Rows.Count > 0)
            {
                string errInfo = "";
                try
                {
                    switch (importDto.ImportType)
                    {
                        case "Simulation"://排程
                            {
                                string sql2 = "";
                                string needType = "0";
                                string needSource = "";
                                string ctNO = "";
                                string ctName = "";
                                string isAdd_SafeQTY = "0";
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        DataRow dr_tmp = null;
                                        foreach (DataRow _dr2 in XMLElementDataTable.Rows)
                                        {
                                            if (_dr2["是否要轉"].ToString().Trim() == "1")
                                            {
                                                ++re_DATA.TotalCount;
                                                #region 欄位檢查
                                                if (_dr2["客戶編號"].ToString().Trim() != "")
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select CTNO FROM SoftNetMainDB.[dbo].[CTData] where ServerId='{_Fun.Config.ServerId}' and CTNO='{_dr2["客戶編號"].ToString().Trim()}'");
                                                    if (dr_tmp == null) { ++re_DATA.FailCount; continue; }
                                                }
                                                if (_dr2["適用行事曆編號"].ToString().Trim() != "")
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select CalendarName FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' and CalendarName='{_dr2["適用行事曆編號"].ToString().Trim()}'");
                                                    if (dr_tmp == null) { ++re_DATA.FailCount; continue; }
                                                }
                                                if (_dr2["物料編號"].ToString().Trim() != "")
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select PartNO FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{_dr2["物料編號"].ToString().Trim()}'");
                                                    if (dr_tmp == null) { ++re_DATA.FailCount; continue; }
                                                }
                                                if (_dr2["適用生產製程"].ToString().Trim() != "")
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select PP_Name FROM SoftNetSYSDB.[dbo].[PP_ProductProcess] where ServerId='{_Fun.Config.ServerId}' and PP_Name='{_dr2["適用生產製程"].ToString().Trim()}'");
                                                    if (dr_tmp == null) { ++re_DATA.FailCount; continue; }
                                                }
                                                if (_dr2["適用工廠編號"].ToString().Trim() != "")
                                                {
                                                    dr_tmp = db.DB_GetFirstDataByDataRow($"select FactoryName FROM SoftNetMainDB.[dbo].[Factory] where ServerId='{_Fun.Config.ServerId}' and FactoryName='{_dr2["適用工廠編號"].ToString().Trim()}'");
                                                    if (dr_tmp == null) { ++re_DATA.FailCount; continue; }
                                                }
                                                #endregion

                                                if (_dr2["訂單編號"].ToString().Trim() != "") { needType = "1"; needSource = _dr2["訂單編號"].ToString(); }
                                                else if (_dr2["客戶編號"].ToString().Trim() != "" || _dr2["客戶名稱"].ToString().Trim() != "")
                                                {
                                                    needType = "2"; needSource = _dr2["客戶編號"].ToString();
                                                    if (_dr2["客戶名稱"].ToString().Trim() != "") { needSource = _dr2["客戶名稱"].ToString().Trim(); }
                                                    else if (_dr2["客戶編號"].ToString().Trim() != "") { needSource = _dr2["客戶編號"].ToString().Trim(); }
                                                }
                                                else if (_dr2["是否為補足安全量"].ToString().Trim() != "") { needType = "4"; }
                                                else { needType = "4"; needSource = ""; }
                                                if (_dr2["是否為補足安全量"].ToString().Trim() == "1") { isAdd_SafeQTY = "1"; }
                                                if (_dr2["物料編號"].ToString().Trim() == "") { continue; }
                                                sql2 = $"INSERT INTO SoftNetSYSDB.[dbo].[APS_NeedData] (ServerId,Id,IsAdd_SafeQTY,NeedType,NeedSource,NeedDate,PartNO,NeedQTY,BufferTime,CalendarName,BOMId,Apply_PP_Name,CTNO,CTName,FactoryName) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('X')}','{isAdd_SafeQTY}','{needType}','{needSource}','{_dr2["預計需求日"].ToString()}','{_dr2["物料編號"].ToString()}',{_dr2["需求量"].ToString()},{_dr2["需求日容許正負差(小時)"].ToString()},'{_dr2["適用行事曆編號"].ToString()}','{_dr2["適用BOM編號"].ToString()}','{_dr2["適用生產製程"].ToString()}','{_dr2["客戶編號"].ToString()}','{_dr2["客戶名稱"].ToString()}','{_dr2["適用工廠編號"].ToString()}')";
                                                if (db.DB_SetData(sql2)) { ++re_DATA.OkCount; }
                                            }
                                            //await _Db.ExecSqlAsync(sql2);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                }
                            }
                            break;
                        case "Material"://料件檔
                            {

                                string sql2 = "";
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        foreach (DataRow _dr2 in XMLElementDataTable.Rows)
                                        {
                                            if (_dr2.IsNull("產品編號") || _dr2["產品編號"].ToString().Trim() == "") { continue; }
                                            sql2 = $"INSERT INTO [dbo].[Material] (ServerId,PartNO,PartName,Specification,Class,Unit,FactoryId) VALUES ('{_Fun.Config.ServerId}','{_dr2["產品編號"].ToString()}','{_dr2["產品品名"].ToString()}','{_dr2["產品規格"].ToString()}','5','PCS','Z01187W3K3HE')";
                                            //db.DB_SetData(sql2);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errInfo += $"<br />{ex.Message}";
                                }
                            }
                            break;
                        case "PPName_AND_BOM"://獨立製程與BOM
                            {
                                string sql2 = "";
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        char re_Type = '0';
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mBOMPartName = "";
                                        string mBOMId = "";

                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("BOM與製程流程") && _dr2["BOM與製程流程"].ToString().Trim() != "" && _dr2["BOM與製程流程"].ToString().Trim() == "母件名稱：")
                                                {
                                                    if (mBOMPartNO != "")
                                                    {
                                                        if (db.DB_GetQueryCount($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where PartNO='{mBOMPartNO}' and Main_Item='1'") <= 0)
                                                        {
                                                            int j = 0;
                                                            string stationNO = "";
                                                            string outPackType = "";
                                                            string mFNO = "";
                                                            string isEnd = "0";
                                                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{mBOMPartNO}'");
                                                            if (tmp == null) { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO) VALUES ('{_Fun.Config.ServerId}','{mBOMPartNO}','{mBOMPartName}','4','PCS','a1')"); }
                                                            mBOMId = _Str.NewId('Z');
                                                            if (pp[(pp.Count - 1)][0] == "") { stationNO = pp[(pp.Count - 1)][2]; mFNO = ""; outPackType = "0"; } else { stationNO = _Fun.Config.OutPackStationName; mFNO = pp[(pp.Count - 1)][0]; outPackType = "1"; }
                                                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[ServerId],[PartNO],[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],Station_DIS_Remark,OutPackType) VALUES
                                                                ('{mBOMId}','{_Fun.Config.ServerId}','{mBOMPartNO}','1','{DateTime.Now.ToString("yyyy/MM/dd")}','{DateTime.Now.AddYears(10).ToString("yyyy/MM/dd")}','1.000','{mBOMPartNO}_加工製程','{stationNO.Split(',')[0]}','0',{pp.Count.ToString()},'','{pp[(pp.Count - 1)][3]}','{pp[(pp.Count - 1)][1]}','{pp[(pp.Count - 1)][4]}')");
                                                            //if (bomII.Count > 0)
                                                            //{
                                                            db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,Id,BOMId,sn,PartNO,BOMQTY,Class) VALUES
                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{mBOMId}',1,'{mBOMPartNO}',1,'4')");

                                                            //}
                                                            db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess] (ServerId,PP_Name,UpdateTime,CalendarName,FactoryName,LineName) VALUES
                                                                ('{_Fun.Config.ServerId}','{mBOMPartNO}_加工製程','{logDate}','{_Fun.Config.DefaultCalendarName}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}')");

                                                            for (int i = 1; i < pp.Count; i++)
                                                            {
                                                                string classType = "4";
                                                                string[] s = pp[(i - 1)];
                                                                mBOMId = _Str.NewId('Z');
                                                                if (i == 1) { isEnd = "1"; classType = "1"; } else { isEnd = "0"; classType = "4"; }
                                                                if (s[0] == "") { stationNO = s[2]; mFNO = ""; outPackType = "0"; } else { stationNO = _Fun.Config.OutPackStationName; mFNO = s[0]; outPackType = "1"; }

                                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[ServerId],[PartNO],[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],Station_DIS_Remark,OutPackType) VALUES
                                                                ('{mBOMId}','{_Fun.Config.ServerId}','{mBOMPartNO}','0','{DateTime.Now.ToString("yyyy/MM/dd")}','{DateTime.Now.AddYears(10).ToString("yyyy/MM/dd")}','','{mBOMPartNO}_加工製程','{stationNO.Split(',')[0]}','{isEnd}',{i.ToString()},'','{s[3]}','{s[1]}','{s[4]}')");
                                                                string partNO = mBOMPartNO;
                                                                if (bomII.Count > 0)
                                                                {
                                                                    if (i == 1) { partNO = bomII[0]; }
                                                                }
                                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,Id,BOMId,sn,PartNO,BOMQTY,Class) VALUES
                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{mBOMId}',1,'{partNO}',1,'{classType}')");

                                                                if (stationNO.Split(',').Length > 1)
                                                                {
                                                                    foreach (string x in stationNO.Split(','))
                                                                    {
                                                                        j += 1;
                                                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey) VALUES
                                                                ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{mBOMPartNO}_加工製程',{(j).ToString()},{(i).ToString()},'1','{x}','{s[3]}','','0','0')");
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    j += 1;
                                                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey) VALUES
                                                                ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{mBOMPartNO}_加工製程',{(j).ToString()},{(i).ToString()},'0','{stationNO}','{s[3]}','{mFNO}','{outPackType}','0')");
                                                                }
                                                            }
                                                            if (pp[(pp.Count - 1)][0] == "") { stationNO = pp[(pp.Count - 1)][2]; mFNO = ""; outPackType = "0"; } else { stationNO = _Fun.Config.OutPackStationName; mFNO = pp[(pp.Count - 1)][0]; outPackType = "1"; }
                                                            if (pp[(pp.Count - 1)][2].Split(',').Length > 1)
                                                            {
                                                                foreach (string x in stationNO.Split(','))
                                                                {
                                                                    j += 1;
                                                                    if (db.DB_GetQueryCount($"SELECT StationNO FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{x}'") <= 0) { db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','無{x}工站')"); }
                                                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey) VALUES
                                                                ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{mBOMPartNO}_加工製程',{(j).ToString()},{(pp.Count).ToString()},'1','{x}','{pp[(pp.Count - 1)][3]}','','0','0')");
                                                                }
                                                            }
                                                            else
                                                            {
                                                                j += 1;
                                                                if (db.DB_GetQueryCount($"SELECT StationNO FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'") <= 0) { db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','無{stationNO}工站')"); }

                                                                db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey) VALUES
                                                                ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{mBOMPartNO}_加工製程',{(j).ToString()},{(pp.Count).ToString()},'0','{stationNO}','{pp[(pp.Count - 1)][3]}','{mFNO}','{outPackType}','0')");
                                                            }
                                                        }
                                                        bomII.Clear();
                                                        pp.Clear();
                                                    }
                                                    mBOMPartNO = _dr2["供應商名稱"].ToString().Trim();
                                                    mBOMPartName = _dr2["備註"].ToString().Trim();
                                                }
                                                else if (!_dr2.IsNull("BOM與製程流程") && _dr2["BOM與製程流程"].ToString().Trim() != "" && _dr2["BOM與製程流程"].ToString().Trim() == "半成品編號：")
                                                {
                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{_dr2["供應商名稱"].ToString().Trim()}'");
                                                    if (tmp == null) { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO) VALUES ('{_Fun.Config.ServerId}','{_dr2["供應商名稱"].ToString().Trim()}','{_dr2["備註"].ToString().Trim()}','2','PCS','a1')"); }
                                                    bomII.Add(_dr2["供應商名稱"].ToString().Trim());
                                                }
                                                else if (!_dr2.IsNull("BOM與製程流程") && _dr2["BOM與製程流程"].ToString().Trim() != "")
                                                {
                                                    if (_dr2["供應商名稱"].ToString().Trim() == "大正科技機械")
                                                    {
                                                        pp.Add(new string[] { _dr2["供應商編號"].ToString().Trim(), _dr2["備註"].ToString().Trim(), _dr2["適用工站"].ToString().Trim(), _dr2["BOM與製程流程"].ToString().Trim(), "0" });
                                                    }
                                                    else
                                                    {
                                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}' and MFNO='{_dr2["供應商編號"].ToString().Trim()}'");
                                                        if (tmp == null)
                                                        {
                                                            string sfName = "";
                                                            if (_dr2["供應商名稱"].ToString().Trim() != "")
                                                            {
                                                                if (_dr2["供應商名稱"].ToString().Trim().Length >= 4) { sfName = _dr2["供應商名稱"].ToString().Trim().Substring(0, 4); }
                                                                else { sfName = _dr2["供應商名稱"].ToString().Trim(); }
                                                            }
                                                            db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[MFData] (ServerId,MFNO,MFName,SName) VALUES ('{_Fun.Config.ServerId}','{_dr2["供應商編號"].ToString().Trim()}','{_dr2["供應商名稱"].ToString().Trim()}','{sfName}')");
                                                        }
                                                        pp.Add(new string[] { _dr2["供應商編號"].ToString().Trim(), _dr2["備註"].ToString().Trim(), _Fun.Config.OutPackStationName, _dr2["BOM與製程流程"].ToString().Trim(), "1" });

                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                for (int ee = ++tt; ee < XMLElementDataTable.Rows.Count; ee++)
                                                {
                                                    DataRow _dree = XMLElementDataTable.Rows[ee];
                                                    if (!_dree.IsNull("BOM與製程流程") && _dree["BOM與製程流程"].ToString().Trim() != "" && _dree["BOM與製程流程"].ToString().Trim() == "母件名稱：")
                                                    { tt = (ee - 1); }
                                                }
                                                mBOMPartNO = "";
                                                mBOMPartName = "";
                                                mBOMId = "";
                                                bomII.Clear();
                                                pp.Clear();
                                                continue;
                                            }

                                        }
                                        if (mBOMPartNO != "")
                                        {
                                            try
                                            {
                                                if (db.DB_GetQueryCount($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where PartNO='{mBOMPartNO}' and Main_Item='1'") <= 0)
                                                {
                                                    int j = 0;
                                                    string stationNO = "";
                                                    string outPackType = "";
                                                    string mFNO = "";
                                                    string isEnd = "0";
                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{mBOMPartNO}'");
                                                    if (tmp == null) { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO) VALUES ('{_Fun.Config.ServerId}','{mBOMPartNO}','{mBOMPartName}','4','PCS','a1')"); }
                                                    mBOMId = _Str.NewId('Z');
                                                    if (pp[(pp.Count - 1)][0] == "") { stationNO = pp[(pp.Count - 1)][2]; mFNO = ""; outPackType = "0"; } else { stationNO = _Fun.Config.OutPackStationName; mFNO = pp[(pp.Count - 1)][0]; outPackType = "1"; }
                                                    db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[ServerId],[PartNO],[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],Station_DIS_Remark,OutPackType) VALUES
                                                                ('{mBOMId}','{_Fun.Config.ServerId}','{mBOMPartNO}','1','{DateTime.Now.ToString("yyyy/MM/dd")}','{DateTime.Now.AddYears(10).ToString("yyyy/MM/dd")}','1.000','{mBOMPartNO}_加工製程','{stationNO.Split(',')[0]}','0',{pp.Count.ToString()},'','{pp[(pp.Count - 1)][3]}','{pp[(pp.Count - 1)][1]}','{pp[(pp.Count - 1)][4]}')");
                                                    //if (bomII.Count > 0)
                                                    //{
                                                    db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,Id,BOMId,sn,PartNO,BOMQTY,Class) VALUES
                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{mBOMId}',1,'{mBOMPartNO}',1,'4')");

                                                    //}
                                                    db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess] (ServerId,PP_Name,UpdateTime,CalendarName,FactoryName,LineName) VALUES
                                                                ('{_Fun.Config.ServerId}','{mBOMPartNO}_加工製程','{logDate}','{_Fun.Config.DefaultCalendarName}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}')");

                                                    for (int i = 1; i < pp.Count; i++)
                                                    {
                                                        string classType = "4";
                                                        string[] s = pp[(i - 1)];
                                                        mBOMId = _Str.NewId('Z');
                                                        if (i == 1) { isEnd = "1"; classType = "1"; } else { isEnd = "0"; classType = "4"; }
                                                        if (s[0] == "") { stationNO = s[2]; mFNO = ""; outPackType = "0"; } else { stationNO = _Fun.Config.OutPackStationName; mFNO = s[0]; outPackType = "1"; }

                                                        db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[ServerId],[PartNO],[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],Station_DIS_Remark,OutPackType) VALUES
                                                                ('{mBOMId}','{_Fun.Config.ServerId}','{mBOMPartNO}','0','{DateTime.Now.ToString("yyyy/MM/dd")}','{DateTime.Now.AddYears(10).ToString("yyyy/MM/dd")}','','{mBOMPartNO}_加工製程','{stationNO.Split(',')[0]}','{isEnd}',{i.ToString()},'','{s[3]}','{s[1]}','{s[4]}')");
                                                        string partNO = mBOMPartNO;
                                                        if (bomII.Count > 0)
                                                        {
                                                            if (i == 1) { partNO = bomII[0]; }
                                                        }
                                                        db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,Id,BOMId,sn,PartNO,BOMQTY,Class) VALUES
                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{mBOMId}',1,'{partNO}',1,'{classType}')");

                                                        if (stationNO.Split(',').Length > 1)
                                                        {
                                                            foreach (string x in stationNO.Split(','))
                                                            {
                                                                j += 1;
                                                                db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey) VALUES
                                                                ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{mBOMPartNO}_加工製程',{(j).ToString()},{(i).ToString()},'1','{x}','{s[3]}','','0','0')");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            j += 1;
                                                            db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey) VALUES
                                                                ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{mBOMPartNO}_加工製程',{(j).ToString()},{(i).ToString()},'0','{stationNO}','{s[3]}','{mFNO}','{outPackType}','0')");
                                                        }
                                                    }
                                                    if (pp[(pp.Count - 1)][0] == "") { stationNO = pp[(pp.Count - 1)][2]; mFNO = ""; outPackType = "0"; } else { stationNO = _Fun.Config.OutPackStationName; mFNO = pp[(pp.Count - 1)][0]; outPackType = "1"; }
                                                    if (pp[(pp.Count - 1)][2].Split(',').Length > 1)
                                                    {
                                                        foreach (string x in stationNO.Split(','))
                                                        {
                                                            j += 1;
                                                            if (db.DB_GetQueryCount($"SELECT StationNO FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{x}'") <= 0) { db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','無{x}工站')"); }
                                                            db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey) VALUES
                                                                ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{mBOMPartNO}_加工製程',{(j).ToString()},{(pp.Count).ToString()},'1','{x}','{pp[(pp.Count - 1)][3]}','','0','0')");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        j += 1;
                                                        if (db.DB_GetQueryCount($"SELECT StationNO FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}' and StationNO='{stationNO}'") <= 0) { db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','無{stationNO}工站')"); }

                                                        db.DB_SetData($@"INSERT INTO SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] (Id,ServerId,FactoryName,LineName,PP_Name,DisplaySN,IndexSN,IndexSN_Merge,StationNO,DisplayName,MFNO,OutPackType,SerialNOKey) VALUES
                                                                ('{_Str.NewId('Y')}','{_Fun.Config.ServerId}','{_Fun.Config.DefaultFactoryName}','{_Fun.Config.DefaultLineName}','{mBOMPartNO}_加工製程',{(j).ToString()},{(pp.Count).ToString()},'0','{stationNO}','{pp[(pp.Count - 1)][3]}','{mFNO}','{outPackType}','0')");
                                                    }
                                                }
                                                bomII.Clear();
                                                pp.Clear();
                                            }
                                            catch (Exception ex)
                                            {
                                                db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "萬_PPName_AND_BOM"://成品生產流程.xlsx
                            {
                                string sql2 = "";
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        char re_Type = '0';
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mBOMPartName = "";
                                        string mBOMId = "";

                                        DataRow tmp = null;
                                        int indexNO = 0;
                                        string Main_Item = "0";
                                        string Version = "";
                                        string Apply_StationNO = "";
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];

                                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='02' and PartNO='{_dr2["PartNO"].ToString().Trim()}'");
                                            if (tmp != null)
                                            {
                                                indexNO = 0;
                                                mBOMPartNO = _dr2["PartNO"].ToString().Trim();
                                                Apply_StationNO = "";
                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set Class='4' where ServerId='02' and PartNO='{mBOMPartNO}'");
                                                if (_dr2["第五站"].ToString().Trim() != "") { indexNO = 5; }
                                                else if (_dr2["第四站"].ToString().Trim() != "") { indexNO = 4; }
                                                else if (_dr2["第三站"].ToString().Trim() != "") { indexNO = 3; }
                                                else if (_dr2["第二站"].ToString().Trim() != "") { indexNO = 2; }
                                                else if (_dr2["第一站"].ToString().Trim() != "") { indexNO = 1; }
                                                if (indexNO > 0)
                                                {

                                                    for (int qq = indexNO; qq >= 1; --qq)
                                                    {
                                                        if (qq == 5) { Apply_StationNO = _dr2["第五站"].ToString().Trim(); }
                                                        else if (qq == 4) { Apply_StationNO = _dr2["第四站"].ToString().Trim(); }
                                                        else if (qq == 3) { Apply_StationNO = _dr2["第三站"].ToString().Trim(); }
                                                        else if (qq == 2) { Apply_StationNO = _dr2["第二站"].ToString().Trim(); }
                                                        else if (qq == 1) { Apply_StationNO = _dr2["第一站"].ToString().Trim(); }

                                                        if (qq == indexNO) { Main_Item = "1"; Version = "1.000"; } else { Main_Item = "0"; Version = ""; }
                                                        mBOMId = _Str.NewId('Z');
                                                        //db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOM] ([Id],[ServerId],[PartNO],[Main_Item],[EffectiveDate],[ExpiryDate],[Version],[Apply_PP_Name],[Apply_StationNO],[IsEnd],[IndexSN],[Station_Custom_IndexSN],[StationNO_Custom_DisplayName],Station_DIS_Remark,OutPackType) VALUES
                                                        //        ('{mBOMId}','02','{mBOMPartNO}','{Main_Item}','{DateTime.Now.ToString("yyyy/MM/dd")}','{DateTime.Now.AddYears(10).ToString("yyyy/MM/dd")}','{Version}','{mBOMPartNO}_加工製程','{Apply_StationNO}','0',{qq.ToString()},'','第{qq.ToString()}次加工','','0')");
                                                        //db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[BOMII] (ServerId,Id,BOMId,sn,PartNO,BOMQTY,Class) VALUES
                                                        //        ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{mBOMId}',1,'{mBOMPartNO}',1,'4')");

                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "萬_PPName_AND_BOM_二次"://產品別工令分析表.xlsx
                            {
                                string sql2 = "";
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        char re_Type = '0';
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mBOMPartName = "";
                                        string mBOMId = "";

                                        DataRow tmp = null;
                                        int indexNO = 0;
                                        string Main_Item = "0";
                                        string Version = "";
                                        string Apply_StationNO = "";
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            //粗胚料號   產品編號
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];

                                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='02' and PartNO='{_dr2["產品編號"].ToString().Trim()}'");
                                            if (tmp != null)
                                            {
                                                tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='02' and PartNO='{_dr2["粗胚料號"].ToString().Trim()}'");
                                                if (tmp != null)
                                                {
                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[BOM] where ServerId='02' and PartNO='{_dr2["產品編號"].ToString().Trim()}' and IsEnd='1'");
                                                    if (tmp != null)
                                                    {
                                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[BOMII] set PartNO='{_dr2["粗胚料號"].ToString().Trim()}' where ServerId='02' and BOMId='{tmp["Id"].ToString().Trim()}'");
                                                    }
                                                    else
                                                    {
                                                        string _s = $"SELECT * from SoftNetMainDB.[dbo].[BOM] where ServerId='02' and PartNO='{_dr2["產品編號"].ToString().Trim()}' and IsEnd='1'";
                                                        string _s2 = "";
                                                    }
                                                }

                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "萬_料號"://品名料號對照表.xlsx
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        char re_Type = '0';
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "";
                                        string mRemark = "";
                                        string mPartNOName = "";
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("C") && _dr2["C"].ToString().Trim() != "" && _dr2["C"].ToString().Trim() == "PC")
                                                {
                                                    mBOMPartNO = _dr2["A"].ToString().Trim();
                                                    if (mBOMPartNO != "" && db.DB_GetQueryCount($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='02' and PartNO='{mBOMPartNO}'") <= 0)
                                                    {
                                                        mClass = _dr2["E"].ToString().Trim();
                                                        if (mClass == "外購") { mClass = "2"; }
                                                        else if (mClass == "外加工") { mClass = "3"; }
                                                        else if (mClass == "自製")
                                                        {
                                                            if (_dr2["D"].ToString().Trim() == "成品") { mClass = "5"; }
                                                            else { mClass = "4"; }
                                                        }
                                                        else
                                                        { mClass = "1"; }
                                                        if (_dr2["I"].ToString().Trim() == "") { mRemark = "NULL"; } else { mRemark = $"'{_dr2["I"].ToString().Trim()}'"; }
                                                        mPartNOName = _dr2["B"].ToString().Trim().Replace("'", "\"");
                                                        db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO,Remark) VALUES ('02','{mBOMPartNO}','{mPartNOName}','{mClass}','PCS','a1',{mRemark})");
                                                    }
                                                    else { mBOMPartNO = ""; }
                                                }
                                                else if (_dr2["B"].ToString().Trim() != "" && _dr2["A"].ToString().Trim() == "" && _dr2["C"].ToString().Trim() == "" && _dr2["D"].ToString().Trim() == "" && _dr2["E"].ToString().Trim() == "" && _dr2["F"].ToString().Trim() == "")
                                                {
                                                    if (mBOMPartNO != "")
                                                    {
                                                        mPartNOName = _dr2["B"].ToString().Trim().Replace("'", "\"");
                                                        db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set Specification='{mPartNOName}' where ServerId='02' and PartNO='{mBOMPartNO}'");
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正成品BOM"://大正BOM結構.xlsx
                            List<string> errLog = new List<string>();
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        List<string[]> bomII = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mBOM_ID = "";
                                        bool is_w = false;
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("A") && _dr2["A"].ToString().IndexOf("產品母件：") == 0)
                                                {
                                                    string ddd = _dr2["A"].ToString().Replace("產品母件：", "").Trim();
                                                    ddd = ddd.Substring(0, (ddd.IndexOf(" ") + 1)).Trim();
                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{ddd}'");
                                                    if (tmp == null) { errLog.Add(ddd); continue; }
                                                    if (ddd == mBOMPartNO && is_w) { continue; }
                                                    is_w = false;
                                                    if (mBOM_ID !="" &&bomII.Count > 0)
                                                    {
                                                        //處裡bomII,並清除
                                                        int tmp_i = 0;
                                                        db.DB_SetData($"delete FROM [SoftNetMainDB].[dbo].[BOMII] where BOMId='{mBOM_ID}'");
                                                        foreach (string[] s in bomII)
                                                        {
                                                            db.DB_SetData($@"INSERT INTO [SoftNetMainDB].[dbo].[BOMII] ([ServerId],[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class]) VALUES
                                                            ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{mBOM_ID}',{(++tmp_i).ToString()},'{s[0]}',{s[2]},'{s[1]}')");
                                                        }
                                                        bomII.Clear();
                                                    }

                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where ServerId='01' and PartNO='{ddd}' and IsEnd='1'");
                                                    if (tmp == null) { errLog.Add(ddd); mBOMPartNO = ""; mBOM_ID = ""; continue; }
                                                    else
                                                    {
                                                        mBOM_ID = tmp["Id"].ToString();
                                                        mBOMPartNO = tmp["PartNO"].ToString();
                                                    }
                                                }
                                                if (mBOMPartNO != "" && _dr2["A"].ToString().Trim() == "標準設計批量：")
                                                {
                                                    is_w = false; continue;
                                                }
                                                if (mBOMPartNO!="" && _dr2["A"].ToString().Trim()== "項次")
                                                {
                                                    is_w = true;continue;
                                                }

                                                if (is_w && !_dr2.IsNull("C") && _dr2["C"].ToString().Trim()!="")
                                                {
                                                    //寫bomII
                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{_dr2["C"].ToString().Trim()}'");
                                                    if (tmp == null) { errLog.Add(_dr2["C"].ToString().Trim()); continue; }
                                                    else
                                                    {
                                                        if (!_dr2.IsNull("G") && _dr2["G"].ToString().Trim() != "")
                                                        {
                                                            int tmp_i = 0;
                                                            if (int.TryParse(_dr2["G"].ToString(), out tmp_i))
                                                            { bomII.Add(new string[] { tmp["PartNO"].ToString(), tmp["Class"].ToString(), tmp_i.ToString() }); }
                                                            else { bomII.Add(new string[] { tmp["PartNO"].ToString(), tmp["Class"].ToString(), "1" }); }
                                                        }
                                                        else { bomII.Add(new string[] { tmp["PartNO"].ToString(), tmp["Class"].ToString(), "1" }); }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正成品BOM_第二版"://20250321/BOM.xlsx
                            List<string> errLog2 = new List<string>();
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        List<string[]> bomII = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mBOM_ID = "";
                                        bool is_w = false;
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("A") && _dr2["A"].ToString().IndexOf("產品母件：") == 0)
                                                {
                                                    string ddd = _dr2["A"].ToString().Replace("產品母件：", "").Trim();
                                                    ddd = ddd.Substring(0, (ddd.IndexOf(" ") + 1)).Trim();
                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where ServerId='01' and PartNO='{ddd}' and IsEnd='1'");
                                                    if (tmp == null) { errLog2.Add(ddd); continue; }
                                                    if (ddd == mBOMPartNO && is_w) { continue; }
                                                    is_w = false;
                                                    if (mBOM_ID != "" && bomII.Count > 0)
                                                    {
                                                        //處裡bomII,並清除
                                                        int tmp_i = 0;
                                                        db.DB_SetData($"delete FROM [SoftNetMainDB].[dbo].[BOMII] where BOMId='{mBOM_ID}'");
                                                        foreach (string[] s in bomII)
                                                        {
                                                            db.DB_SetData($@"INSERT INTO [SoftNetMainDB].[dbo].[BOMII] ([ServerId],[Id],[BOMId],[sn],[PartNO],[BOMQTY],[Class]) VALUES
                                                            ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{mBOM_ID}',{(++tmp_i).ToString()},'{s[0]}',{s[2]},'{s[1]}')");
                                                        }
                                                        bomII.Clear();
                                                    }

                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[BOM] where ServerId='01' and PartNO='{ddd}' and IsEnd='1'");
                                                    if (tmp == null) { errLog2.Add(ddd); mBOMPartNO = ""; mBOM_ID = ""; continue; }
                                                    else
                                                    {
                                                        mBOM_ID = tmp["Id"].ToString();
                                                        mBOMPartNO = tmp["PartNO"].ToString();
                                                    }
                                                }
                                                if (mBOMPartNO != "" && _dr2["A"].ToString().Trim() == "標準設計批量：")
                                                {
                                                    is_w = false; continue;
                                                }
                                                if (mBOMPartNO != "" && _dr2["A"].ToString().Trim() == "項次")
                                                {
                                                    is_w = true; continue;
                                                }

                                                if (is_w && !_dr2.IsNull("B") && _dr2["B"].ToString().Trim() != "")
                                                {
                                                    //寫bomII
                                                    tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{_dr2["B"].ToString().Trim()}'");
                                                    if (tmp == null) { errLog2.Add(_dr2["B"].ToString().Trim()); continue; }
                                                    else
                                                    {
                                                        if (!_dr2.IsNull("C") && _dr2["C"].ToString().Trim() != "")
                                                        {
                                                            int tmp_i = 0;
                                                            if (int.TryParse(_dr2["C"].ToString(), out tmp_i))
                                                            { bomII.Add(new string[] { tmp["PartNO"].ToString(), tmp["Class"].ToString(), tmp_i.ToString() }); }
                                                            else { bomII.Add(new string[] { tmp["PartNO"].ToString(), tmp["Class"].ToString(), "1" }); }
                                                        }
                                                        else { bomII.Add(new string[] { tmp["PartNO"].ToString(), tmp["Class"].ToString(), "1" }); }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正_原物料"://20250321/物料型號.xlsx
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        char re_Type = '0';
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "2";
                                        string mRemark = "";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mBOMPartNO = _dr2["產品編號"].ToString().Trim();
                                                    mPartNOName = _dr2["產品品名"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    mSpecification = _dr2["產品規格"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    string ppp = mBOMPartNO.Substring((mBOMPartNO.Length - 2), 2);
                                                    if (mBOMPartNO != "")
                                                    {
                                                        if (db.DB_GetQueryCount($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{mBOMPartNO}'") <= 0)
                                                        {
                                                            if (mBOMPartNO.Substring((mBOMPartNO.Length - 2), 2) == "-N")
                                                            {
                                                                mClass = "2";
                                                            }
                                                            else
                                                            {
                                                                mClass = "4";
                                                            }
                                                            if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO,Specification) VALUES ('01','{mBOMPartNO}','{mPartNOName}','{mClass}','PCS','a1','{mSpecification}')"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (!db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set PartName='{mPartNOName}',Specification='{mSpecification}' where ServerId='01' and PartNO='{mBOMPartNO}'"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正_成品料"://20250321/產品型號.xlsx
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        char re_Type = '0';
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "2";
                                        string mRemark = "";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mBOMPartNO = _dr2["產品編號"].ToString().Trim();
                                                    mPartNOName = _dr2["產品品名"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    mSpecification = _dr2["產品規格"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    if (mBOMPartNO != "")
                                                    {
                                                        if (db.DB_GetQueryCount($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{mBOMPartNO}'") <= 0)
                                                        {
                                                            if (mBOMPartNO.Substring((mBOMPartNO.Length - 2), 2) == "-N")
                                                            {
                                                                mClass = "2";
                                                            }
                                                            else
                                                            {
                                                                mClass = "5";
                                                            }
                                                            if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO,Specification) VALUES ('01','{mBOMPartNO}','{mPartNOName}','{mClass}','PCS','a1','{mSpecification}')"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (!db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set PartName='{mPartNOName}',Specification='{mSpecification}' where ServerId='01' and PartNO='{mBOMPartNO}'"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正_機械倉_關聯料號"://20250321/智倉零件.xlsx
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        char re_Type = '0';
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "2";
                                        string mRemark = "";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("A") && _dr2["A"].ToString().Trim() != "")
                                                {
                                                    mBOMPartNO = _dr2["A"].ToString().Trim();
                                                    if (mBOMPartNO != "")
                                                    {
                                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{mBOMPartNO}'");
                                                        if (tmp!=null)
                                                        {
                                                            db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set APS_Default_StoreNO='a2',APS_Default_StoreSpacesNO='' where ServerId='01' and PartNO='{mBOMPartNO}'");
                                                            if (db.DB_GetQueryCount($"SELECT Id from [SoftNetMainDB].[dbo].[TotalStock] where ServerId='01' and PartNO='{mBOMPartNO}'") > 0)
                                                            { db.DB_SetData($"update [SoftNetMainDB].[dbo].[TotalStock] set StoreNO='a2',StoreSpacesNO='' where ServerId='01' and PartNO='{mBOMPartNO}'"); }
                                                            else
                                                            {
                                                                db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] ([ServerId],[Id],[StoreNO],[StoreSpacesNO],[PartNO] ,[QTY],[WeightQty]) VALUES
                                                                ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','a2','','{mBOMPartNO}',0,0)");
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正_電子燈_關聯料號"://20250321/mes編號表.xlsx
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        char re_Type = '0';
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "2";
                                        string mRemark = "";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        DataRow tmp = null;
                                        db.DB_SetData($"delete from [SoftNetMainDB].[dbo].[StoreII] where ServerId='01' and StoreNO='a1'");
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("指示燈編號") && _dr2["指示燈編號"].ToString().Trim() != "" && !_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mBOMPartNO = _dr2["產品編號"].ToString().Trim();
                                                    mRemark= _dr2["指示燈編號"].ToString().Trim();
                                                    if (mBOMPartNO != "")
                                                    {
                                                        db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[StoreII] ([ServerId],[Id],[StoreNO],[StoreSpacesNO],[StoreSpacesName],[Remark],[Config_macID]) VALUES
                                                                ('01','{_Str.NewId('Z')}','a1','{mRemark}','','','{mRemark}')");

                                                        foreach (string s in mBOMPartNO.Split(";"))
                                                        {
                                                            if (s == "") { continue; }
                                                            tmp = db.DB_GetFirstDataByDataRow($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{s}'");
                                                            if (tmp != null)
                                                            {
                                                                db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set APS_Default_StoreNO='a1',APS_Default_StoreSpacesNO='{mRemark}' where ServerId='01' and PartNO='{s}'");
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正_半成品_3":
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "4";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mBOMPartNO = _dr2["產品編號"].ToString().Trim();
                                                    mPartNOName = _dr2["產品品名"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    mSpecification = _dr2["產品規格"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    if (mBOMPartNO != "")
                                                    {
                                                        if (db.DB_GetQueryCount($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{mBOMPartNO}'") <= 0)
                                                        {

                                                            if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO,Specification) VALUES ('01','{mBOMPartNO}','{mPartNOName}','{mClass}','PCS','a1','{mSpecification}')"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (!db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set Class='4',APS_Default_StoreNO='a1',PartName='{mPartNOName}',Specification='{mSpecification}' where ServerId='01' and PartNO='{mBOMPartNO}'"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }

                            break;
                        case "大正_成品_3":
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "5";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mBOMPartNO = _dr2["產品編號"].ToString().Trim();
                                                    mPartNOName = _dr2["產品品名"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    mSpecification = _dr2["產品規格"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    if (mBOMPartNO != "")
                                                    {
                                                        if (db.DB_GetQueryCount($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{mBOMPartNO}'") <= 0)
                                                        {

                                                            if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO,Specification) VALUES ('01','{mBOMPartNO}','{mPartNOName}','{mClass}','PCS','b001','{mSpecification}')"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (!db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set Class='5',APS_Default_StoreNO='b001',PartName='{mPartNOName}',Specification='{mSpecification}' where ServerId='01' and PartNO='{mBOMPartNO}'"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }

                            break;
                        case "大正_素材_3":
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "1";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        string mfno = "";
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mfno = _dr2["供應商"].ToString().Trim();
                                                    int dfg = mfno.Length;
                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[MFData] where ServerId='01' and  (MFName like'%{_dr2["供應商"].ToString().Trim()}%' or SName like'%{_dr2["供應商"].ToString().Trim()}%')");
                                                    if (tmp == null)
                                                    {
                                                        mfno = "";
                                                        if (!bomII.Contains(_dr2["供應商"].ToString().Trim()))
                                                        {
                                                            bomII.Add(_dr2["供應商"].ToString().Trim());
                                                            re_DATA.ErrorMsg = $"{re_DATA.ErrorMsg}<br />{_dr2["供應商"].ToString().Trim()}";
                                                        }
                                                    }
                                                    else { mfno = tmp["MFNO"].ToString(); }
                                                    mBOMPartNO = _dr2["產品編號"].ToString().Trim();
                                                    mPartNOName = _dr2["產品品名"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    mSpecification = _dr2["產品規格"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    if (mBOMPartNO != "")
                                                    {
                                                        if (db.DB_GetQueryCount($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{mBOMPartNO}'") <= 0)
                                                        {

                                                            if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO,Specification,APS_Default_MFNO) VALUES ('01','{mBOMPartNO}','{mPartNOName}','{mClass}','PCS','b001','{mSpecification}','{mfno}')"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (!db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set  Class='1',APS_Default_StoreNO='b001',APS_Default_MFNO='{mfno}',PartName='{mPartNOName}',Specification='{mSpecification}' where ServerId='01' and PartNO='{mBOMPartNO}'"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                        string sdf = "";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正_採購半成品_3":
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "2";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        string mfno = "";
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mfno = _dr2["供應商"].ToString().Trim();
                                                    int dfg = mfno.Length;
                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[MFData] where ServerId='01' and  (MFName like'%{_dr2["供應商"].ToString().Trim()}%' or SName like'%{_dr2["供應商"].ToString().Trim()}%')");
                                                    if (tmp == null)
                                                    {
                                                        mfno = "";
                                                        if (!bomII.Contains(_dr2["供應商"].ToString().Trim()))
                                                        {
                                                            bomII.Add(_dr2["供應商"].ToString().Trim());
                                                            re_DATA.ErrorMsg = $"{re_DATA.ErrorMsg}<br />{_dr2["供應商"].ToString().Trim()}";
                                                        }
                                                    }
                                                    else { mfno = tmp["MFNO"].ToString(); }
                                                    mBOMPartNO = _dr2["產品編號"].ToString().Trim();
                                                    mPartNOName = _dr2["產品品名"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    mSpecification = _dr2["產品規格"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    if (mBOMPartNO != "")
                                                    {
                                                        if (db.DB_GetQueryCount($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{mBOMPartNO}'") <= 0)
                                                        {

                                                            if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO,Specification,APS_Default_MFNO) VALUES ('01','{mBOMPartNO}','{mPartNOName}','{mClass}','PCS','b001','{mSpecification}','{mfno}')"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (!db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set  Class='2',APS_Default_StoreNO='b001',APS_Default_MFNO='{mfno}',PartName='{mPartNOName}',Specification='{mSpecification}' where ServerId='01' and PartNO='{mBOMPartNO}'"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                        string sdf = "";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }

                            break;
                        case "大正_採購成品_3":
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "2";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        string mfno = "";
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mfno = _dr2["供應商"].ToString().Trim();
                                                    int dfg = mfno.Length;
                                                    tmp = db.DB_GetFirstDataByDataRow($"select * from SoftNetMainDB.[dbo].[MFData] where ServerId='01' and (MFName like'%{_dr2["供應商"].ToString().Trim()}%' or SName like'%{_dr2["供應商"].ToString().Trim()}%')");
                                                    if (tmp == null)
                                                    {
                                                        mfno = "";
                                                        if (!bomII.Contains(_dr2["供應商"].ToString().Trim()))
                                                        {
                                                            bomII.Add(_dr2["供應商"].ToString().Trim());
                                                            re_DATA.ErrorMsg = $"{re_DATA.ErrorMsg}<br />{_dr2["供應商"].ToString().Trim()}";
                                                        }
                                                    }
                                                    else { mfno = tmp["MFNO"].ToString(); }
                                                    mBOMPartNO = _dr2["產品編號"].ToString().Trim();
                                                    mPartNOName = _dr2["產品品名"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    mSpecification = _dr2["產品規格"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    if (mBOMPartNO != "")
                                                    {
                                                        if (db.DB_GetQueryCount($"SELECT [PartNO] FROM SoftNetMainDB.[dbo].[Material] where ServerId='01' and PartNO='{mBOMPartNO}'") <= 0)
                                                        {

                                                            if (!db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Class,Unit,APS_Default_StoreNO,Specification,APS_Default_MFNO) VALUES ('01','{mBOMPartNO}','{mPartNOName}','{mClass}','PCS','b001','{mSpecification}','{mfno}')"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (!db.DB_SetData($"update SoftNetMainDB.[dbo].[Material] set  Class='2',APS_Default_StoreNO='b001',APS_Default_MFNO='{mfno}',PartName='{mPartNOName}',Specification='{mSpecification}' where ServerId='01' and PartNO='{mBOMPartNO}'"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                        string sdf = "";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正_共應商_3":
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mMFNO = "";
                                        string mMFName = "";
                                        string mSName = "";
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mMFNO = _dr2["產品編號"].ToString().Trim();
                                                    mMFName = _dr2["廠商名稱"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    mSName = _dr2["簡稱"].ToString().Trim().Replace("\"", "＂").Replace("'", "’");
                                                    if (mMFNO != "")
                                                    {
                                                        if (db.DB_GetQueryCount($"SELECT [MFNO] FROM SoftNetMainDB.[dbo].[MFData] where ServerId='01' and MFNO='{mMFNO}'") <= 0)
                                                        {

                                                            if (!db.DB_SetData($@"INSERT INTO SoftNetMainDB.[dbo].[MFData] ([ServerId],[MFNO],[MFName],[SName],[UniFormNO],[TEL],[FAX],[ContactMan],[ContactTEL],[EMail],[Address],[Remark])
                                                                                    VALUES ('01','{mMFNO}','{mMFName}','{mSName}','{_dr2["統一編號"].ToString()}','{_dr2["TEL"].ToString()}','{_dr2["FAX"].ToString()}','{_dr2["聯絡人"].ToString()}','{_dr2["聯絡TEL"].ToString()}','{_dr2["EMail"].ToString()}','{_dr2["工廠地址"].ToString()}','{_dr2["發票地址"].ToString()}')"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (!db.DB_SetData($"update SoftNetMainDB.[dbo].[MFData] set MFName='{mMFName}',SName='{mSName}',UniFormNO='{_dr2["統一編號"].ToString()}',TEL='{_dr2["TEL"].ToString()}',FAX='{_dr2["FAX"].ToString()}',ContactMan='{_dr2["聯絡人"].ToString()}',ContactTEL='{_dr2["聯絡TEL"].ToString()}',EMail='{_dr2["EMail"].ToString()}',Address='{_dr2["工廠地址"].ToString()}',Remark='{_dr2["發票地址"].ToString()}' where ServerId='01' and MFNO='{mMFNO}'"))
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }

                            break;
                        case "大正_刀具資料":
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        DataRow dr_tmp = null;
                                        int sno = 1;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("A") && _dr2["A"].ToString().Trim() != "")
                                                {
                                                    //(1-1)刀片/CCMT060204N-AC830P
                                                    string x = "";
                                                    string y = "";
                                                    string PartName = "";
                                                    string Specification = "";
                                                    string StoreSpacesNO = "";
                                                    string partNO = _dr2["A"].ToString().Trim();
                                                    string[] tmp = partNO.Split(')');
                                                    
                                                    if (tmp != null && tmp.Length > 1 && partNO.IndexOf("(") == 0 && partNO.IndexOf(")") > 0)
                                                    {
                                                        string tmp01 = partNO.Substring(1, (partNO.IndexOf(")") - 1)).Trim();
                                                        string tmp02 = partNO.Substring((partNO.IndexOf(")") + 1)).Trim();
                                                        string[] tmp_storII = tmp01.Split('-');
                                                        if (tmp01 != "" && tmp02 != "" && tmp_storII != null && tmp_storII.Length == 2)
                                                        {
                                                            if (tmp02.IndexOf("/") > 0 && tmp02.IndexOf("Φ") > 0 && tmp02.IndexOf("Φ") > tmp02.IndexOf("/"))
                                                            {
                                                                PartName = tmp02.Substring(0, (tmp02.IndexOf("/") - 0)).Trim();
                                                                Specification = tmp02.Substring((tmp02.IndexOf("/") + 1)).Trim();
                                                            }
                                                            else if (tmp02.IndexOf("/") > 0 && tmp02.IndexOf("Φ") > 0 && tmp02.IndexOf("/") > tmp02.IndexOf("Φ"))
                                                            {
                                                                PartName = tmp02.Substring(0, (tmp02.IndexOf("Φ") - 0)).Trim();
                                                                Specification = tmp02.Substring((tmp02.IndexOf("Φ") + 1)).Trim();
                                                            }
                                                            else if (tmp02.IndexOf("/") > 0 && tmp02.IndexOf("Φ") < 0)
                                                            {
                                                                PartName = tmp02.Substring(0, (tmp02.IndexOf("/") - 0)).Trim();
                                                                Specification = tmp02.Substring((tmp02.IndexOf("/") + 1)).Trim();
                                                            }
                                                            else if (tmp02.IndexOf("/") < 0 && tmp02.IndexOf("Φ") > 0)
                                                            {
                                                                PartName = tmp02.Substring(0, (tmp02.IndexOf("Φ") - 0)).Trim();
                                                                Specification = tmp02.Substring((tmp02.IndexOf("Φ") + 1)).Trim();
                                                            }
                                                            else
                                                            { PartName = tmp02; }
                                                            partNO = $"Z{(++sno).ToString().PadLeft(5, '0')}";
                                                            PartName = PartName.Replace("(", "（").Replace(")", "）").Replace("'", "’");
                                                            Specification = Specification.Replace("(", "（").Replace(")", "）").Replace("'", "’");

                                                            StoreSpacesNO = $"1.{tmp_storII[0]}.{tmp_storII[1]}";
                                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[StoreII] where ServerId='{_Fun.Config.ServerId}' and StoreNO='c001' and StoreSpacesNO='{StoreSpacesNO}'");
                                                            if (dr_tmp == null) { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[StoreII] ([ServerId],[Id],[StoreNO],[StoreSpacesNO],[StoreSpacesName],[Remark],[Config_macID]) VALUES ('01','{_Str.NewId('Z')}','c001','{StoreSpacesNO}','','','')"); }
                                                            dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and PartNO='{partNO}'");
                                                            if (dr_tmp == null) { db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[Material] (ServerId,PartNO,PartName,Specification,Class,Unit,APS_Default_StoreNO,APS_Default_StoreSpacesNO) VALUES ('{_Fun.Config.ServerId}','{partNO}','{PartName}','{Specification}','6','PCS','c001','{StoreSpacesNO}')"); }


                                                        }
                                                        string _s = "";
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }


                                        #region 檢查TotalStock是否有資料
                                        DataTable dt = db.DB_GetData($"SELECT * from SoftNetMainDB.[dbo].[Material] where ServerId='{_Fun.Config.ServerId}' and APS_Default_StoreNO='c001'");
                                        if (dt != null && dt.Rows.Count > 0)
                                        {
                                            string tmp_sq = "";
                                            foreach (DataRow dr_II in dt.Rows)
                                            {

                                                if (dr_II["APS_Default_StoreSpacesNO"].ToString() != "") { tmp_sq = $" and StoreNO='{dr_II["APS_Default_StoreNO"].ToString()}' and StoreSpacesNO='{dr_II["APS_Default_StoreSpacesNO"].ToString()}'"; }
                                                else { tmp_sq = $" and StoreNO='{dr_II["APS_Default_StoreNO"].ToString()}'"; }
                                                dr_tmp = db.DB_GetFirstDataByDataRow($"SELECT * from SoftNetMainDB.[dbo].[TotalStock] where ServerId='{_Fun.Config.ServerId}' and PartNO='{dr_II["PartNO"].ToString()}' {tmp_sq}");
                                                if (dr_tmp == null)
                                                {
                                                    if (db.DB_SetData($"INSERT INTO SoftNetMainDB.[dbo].[TotalStock] (ServerId,[Id],[StoreNO],[StoreSpacesNO],[PartNO],[QTY]) VALUES ('{_Fun.Config.ServerId}','{_Str.NewId('Z')}','{dr_II["APS_Default_StoreNO"].ToString()}','{dr_II["APS_Default_StoreSpacesNO"].ToString()}','{dr_II["PartNO"].ToString()}',0)"))
                                                    {
                                                        string _s = "";
                                                    }
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }
                            break;
                        case "大正_庫存數量導入":
                            {
                                try
                                {
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        string logDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        char re_Type = '0';
                                        List<string> bomII = new List<string>();
                                        List<string[]> pp = new List<string[]>();
                                        string mBOMPartNO = "";
                                        string mClass = "2";
                                        string mRemark = "";
                                        string mPartNOName = "";
                                        string mSpecification = "";
                                        int qty = 0;
                                        DataRow tmp = null;
                                        for (int tt = 0; tt < XMLElementDataTable.Rows.Count; tt++)
                                        {
                                            DataRow _dr2 = XMLElementDataTable.Rows[tt];
                                            try
                                            {
                                                if (!_dr2.IsNull("產品編號") && _dr2["產品編號"].ToString().Trim() != "")
                                                {
                                                    mBOMPartNO = _dr2["產品編號"].ToString().Trim();
                                                    if (mBOMPartNO != "" && int.TryParse(_dr2["實際在庫量"].ToString(), out qty))
                                                    {
                                                        tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].Material where ServerId='01' and PartNO='{mBOMPartNO}'");
                                                        if (tmp!=null)
                                                        {
                                                            tmp = db.DB_GetFirstDataByDataRow($"SELECT * FROM SoftNetMainDB.[dbo].[TotalStock] where ServerId='01' and PartNO='{mBOMPartNO}' and StoreNO='{tmp["APS_Default_StoreNO"].ToString()}' and StoreSpacesNO='{tmp["APS_Default_StoreSpacesNO"].ToString()}'");
                                                            if (tmp != null)
                                                            {
                                                                if (!db.DB_SetData($"update SoftNetMainDB.[dbo].[TotalStock] set QTY={qty.ToString()} where Id='{tmp["Id"].ToString()}'"))
                                                                {
                                                                    string _s = "";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                string _s = "";
                                                            }
                                                        }
                                                        else
                                                        {
                                                            string _s = "";
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                //db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('{mBOMPartNO}','{ex.Message.Replace("'", "|")}')");
                                                continue;
                                            }

                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string _s = "";
                                    using (DBADO db = new DBADO("1", _Fun.Config.Db))
                                    {
                                        db.DB_SetData($@"INSERT INTO SoftNetLogDB.[dbo].[test_BOM_INFO] (Log01_PartNO,ExINFO) VALUES ('主程式Error','{ex.Message}')");
                                    }
                                }
                            }

                            break;
                    }
                }
                catch (Exception ex)
                {
                    errInfo += $"<br />{ex.Message}";
                }
                if (errInfo != "")
                { re_DATA.ErrorMsg = $"匯入出現錯誤, 請將下列訊息提供給系統管理者. {errInfo}"; }
                //re_DATA.OkCount = XMLElementDataTable.Rows.Count;
                //re_DATA.TotalCount = XMLElementDataTable.Rows.Count;
                return re_DATA;
            }
            if (XMLElementDataTable == null || XMLElementDataTable.Rows.Count <= 0)
            {
                re_DATA.ErrorMsg = $"無資料可匯入.";
                return re_DATA;
            }



                /*
                }
                else
                {
                    //有傳入excel欄位名稱
                    //check
                    var cellLen = cells.Count();
                    if (cellLen != importDto.ExcelFids.Count)
                    {
                        errorMsg = "importDto.ExcelFids length should be " + cellLen;
                        goto lab_error;
                    }

                    //set colMap
                    for (var i=0; i< cellLen; i++)
                    {
                        var colName = CellXname(cells.ElementAt(i).CellReference);
                        colMap[colName] = i;
                    }
                }
                */

                //initial excelIsDates & set excelFidLen
                var excelIsDates = new List<bool>();        //是否為是日期欄位
            var excelFidLen = excelFids.Count;
            for (var i = 0; i < excelFidLen; i++)
                excelIsDates.Add(false);    //initial
            #endregion

            #region set excelIsDates, modelFids, modelDateFids/Fno/Len, modelNotDateFids/Fno/Len
            /*
            int fno;
            var modelFids = new List<string>();         //全部欄位
            var model = new T();
            foreach (var prop in model.GetType().GetProperties())
            {
                //如果對應的excel欄位不存在, 則不記錄此欄位(skip)
                //var type = prop.GetValue(model, null).GetType();
                var fid = prop.Name;
                fno = excelFids.FindIndex(a => a == fid);
                if (fno < 0)
                    continue;

                modelFids.Add(fid);
                if (prop.PropertyType == typeof(DateTime?))
                    excelIsDates[fno] = true;
            }
            */

            //var modelDateFidLen = modelDateFids.Count;
            //var modelNotDateFidLen = modelNotDateFids.Count;
            #endregion

            #region set fileRows by excel file
            var fileRows = new List<T>();   //excel rows with data(not empty row)
            var excelRowLen = excelRows.LongCount();
            for (var i = importDto.FidRowNo; i < excelRowLen; i++)
            {
                var excelRow = excelRows.ElementAt(i);
                var fileRow = new T();
                /*
                //set datetime column
                //var rowHasCol = false;
                for(var j=0; j<modelDateFidLen; j++)
                {
                    //var cell = cells.ElementAt(modelDateFnos[j]);
                    var cell = excelRow.Descendants<Cell>().ElementAt(j);
                    if (cell.DataType != null)
                    {
                        //rowHasCol = true;
                        value = (cell.DataType == CellValues.SharedString) ? ssTable.ChildElements[int.Parse(cell.CellValue.Text)].InnerText :
                            cell.CellValue.Text;
                        _Model.SetValue(modelRow, modelDateFids[j], DateTime.FromOADate(double.Parse(value)).ToString(rb.uiDtFormat));
                    }
                }
                */

                //write not date column
                //for (var j = 0; j < modelNotDateFidLen; j++)
                //var j = 0;
                foreach (Cell cell in excelRow)
                {
                    /*
                    if (i == 2 && j == 1)
                    {
                        var aa = "aa";
                    }
                    */
                    //var cell = cells.ElementAt(modelNotDateFnos[j]);
                    //var cell = excelRow.Descendants<Cell>().ElementAt(modelNotDateFnos[j]);
                    //colName = ;
                    fno = (int)colMap[CellXname(cell.CellReference)];
                    /*
                    var value = (cell.DataType == CellValues.SharedString) 
                        ? ssTable.ChildElements[int.Parse(cell.CellValue.Text)].InnerText 
                        : cell.CellValue.Text;

                    _Model.SetValue(fileRow, excelFids[fno], excelIsDates[fno]
                        ? DateTime.FromOADate(double.Parse(value)).ToString(uiDtFormat)
                        : value
                    );
                    */
                }

                fileRows.Add(fileRow);
            }
            #endregion
            #endregion

            #region 2.validate fileRows loop
            idx = 0;
            //var error = "";
            foreach (var fileRow in fileRows)
            {
                //validate
                var context = new ValidationContext(fileRow, null, null);
                var results = new List<ValidationResult>();
                if (Validator.TryValidateObject(fileRow, context, results, true))
                {
                    //user validate rule
                    //if (importDto.FnCheckImportRow != null)
                    //    error = importDto.FnCheckImportRow(fileRow);
                    //if (_Str.IsEmpty(error))
                        _okRowNos.Add(idx);
                    //else
                    //    AddError(idx, error);
                }
                else
                {
                    AddErrorByResults(idx, results);
                }
                idx++;
            }
            #endregion

            #region 3.save database for ok rows(call FnSaveImportRows())
            if (_okRowNos.Count > 0)
            {
                //set okRows
                var okRows = new List<T>();
                foreach(var okRowNo in _okRowNos)
                    okRows.Add(fileRows[okRowNo]);

                //call FnSaveImportRows
                idx = 0;
                var saveResults = importDto.FnSaveImportRows(okRows);
                if (saveResults != null)
                {
                    foreach (var result in saveResults)
                    {
                        if (!_Str.IsEmpty(result))
                            AddError(_okRowNos[idx], result);
                        idx++;
                    }
                }
            }
            #endregion

            #region 4.save ok excel file
            //###??? 更換.NET10無法使用
            if (_Str.IsEmpty(importDto.LogRowId))
                importDto.LogRowId = _Str.NewId('O');
            var fileStem = _Str.AddDirSep(dirUpload) + importDto.LogRowId;
            //docx.SaveAs(fileStem + ".xlsx");
            #endregion

            #region 5.save fail excel file (tail _fail.xlsx)
            var failCount = _failRows.Count;
            if (failCount > 0)
            {
                //set excelFnos: excel column map model column
                var excelFnos = new List<int>();
                for (var i = 0; i < excelFidLen; i++)
                {
                    fno = modelFids.FindIndex(a => a == excelFids[i]);
                    excelFnos.Add(fno);    //<0 means no mapping
                }

                //get docx
                var failFilePath = fileStem + "_fail.xlsx";
                File.Copy(importDto.TplPath, failFilePath, true);

                var docx2 = SpreadsheetDocument.Open(failFilePath, true);
                var wbPart2 = docx2.WorkbookPart;
                var wsPart2 = (WorksheetPart)wbPart2.GetPartById(
                    wbPart2.Workbook.Descendants<Sheet>().ElementAt(0).Id);
                var sheetData2 = wsPart2.Worksheet.GetFirstChild<SheetData>();

                var startRow = importDto.FidRowNo;    //insert position
                for (var i = 0; i < failCount; i++)
                {
                    //add row, fill value & copy row style
                    var modelRow = fileRows[_failRows[i].Sn];
                    var newRow = new Row();     //new excel row
                    for (var colNo = 0; colNo < excelFidLen; colNo++)
                    {
                        fno = excelFnos[colNo];
                        var value2 = _Model.GetValue(modelRow, excelFids[colNo]);
                        newRow.Append(new Cell()
                        {
                            CellValue = new CellValue(fno < 0 || value2 == null ? "" : value2.ToString()),
                            DataType = CellValues.String,
                        });
                    }

                    //write cell for error msg
                    newRow.Append(new Cell()
                    {
                        CellValue = new CellValue(_failRows[i].Str),
                        DataType = CellValues.String,
                    });

                    sheetData2.InsertAt(newRow, startRow + i);
                }
                docx2.Save();
                docx2.Dispose();
            }
            #endregion

            #region 6.insert ImportLog table
            var totalCount = fileRows.Count;
            var okCount = totalCount - failCount;
            var sql = $@"
insert into dbo.XpImportLog(Id, Type, FileName,
OkCount, FailCount, TotalCount,
CreatorName, Created)
values('{importDto.LogRowId}', '{importDto.ImportType}', '{fileName}',
{okCount}, {failCount}, {totalCount}, 
'{importDto.CreatorName}', '{_Date.NowDbStr()}')
";
            //await _Db.ExecSqlAsync(sql);
            #endregion

            //7.return import result
            return new ResultImportDto()
            {
                OkCount = okCount,
                FailCount = failCount,
                TotalCount = totalCount,
            };
        }

        private string GetCellValue(SharedStringTable ssTable, Cell cell)
        {
            //SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            var value = cell.CellValue.InnerXml;
            return (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                ? ssTable.ChildElements[Int32.Parse(value)].InnerText
                : value;
        }


        /// <summary>
        /// add row error
        /// </summary>
        /// <param name="error"></param>
        public void AddError(int rowNo, string error)
        {
            if (!_Str.IsEmpty(error))
            {
                _failRows.Add(new SnStrDto()
                {
                    Sn = rowNo,
                    Str = error,
                });
            }
        }

        /// <summary>
        /// add row error for multiple error msg
        /// </summary>
        /// <param name="rowNo"></param>
        /// <param name="results"></param>
        private void AddErrorByResults(int rowNo, List<ValidationResult> results)
        {
            var error = string.Join(RowSep, results.Select(a => a.ErrorMessage).ToList());
            AddError(rowNo, error);
        }

        #region remark code
        /*
        //excel docx to datatable
        private DataTable DocxToDataTable(SpreadsheetDocument docx)
        {
            //open excel
            WorkbookPart wbPart = docx.WorkbookPart;
            SharedStringTablePart ssPart = wbPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringTable ssTable = ssPart.SharedStringTable;

        }

        //驗証資料列是否正確
        public bool ValidDto(T dto)
        {
            //欄位驗証
            var context = new ValidationContext(dto, null, null);
            var results = new List<ValidationResult>();
            if (Validator.TryValidateObject(dto, context, results, true))
                return true;

            //case of error
            AddErrorByResults(results);
            return false;
        }

        private string Init(string asposeLicPath, string uploadFileName, string saveExcelPath, string tplPath, string sheetName)
        {
            //set instance variables
            this.asposeLicPath = asposeLicPath;
            this.uploadFileName = uploadFileName;
            this.saveExcelPath = saveExcelPath;
            this.tplPath = tplPath;
            this.sheetName = sheetName;
            this.sysFileName = Path.GetFileName(saveExcelPath);

            //建立excel connection
            var excelBook = Utils.OpenExcel(asposeLicPath, saveExcelPath);

            //set dataTable _dt
            var error = "";
            this.dt = Utils.ReadWorksheet(excelBook, sheetName, 0, 0, out error, true);
            return error;
        }

        /// <summary>
        /// 驗証excel資料列
        /// </summary>
        /// <returns>status</returns>
        protected void CheckRows()
        {
            //set instance variables
            this.nowIndex = 0;

            //check rows
            foreach (DataRow dr in dt.Rows)
            {
                //reset rowError
                this.rowError = new List<string>();

                //check row
                var model = CheckTableRow(dr);
                if (model != null)
                    okRows.Add(model);

                this.nowIndex++;
            }
        }

        //save error result to excel file
        protected void SaveErrorToExcel(int startRow, int colLen)
        {
            //check
            if (this._errorRows.Count == 0)
                return;

            //copy excel & add _fail for file name
            var fileName = Path.GetDirectoryName(this.saveExcelPath) + _Fun.DirSep
                + Path.GetFileNameWithoutExtension(this.saveExcelPath)
                + "_fail" + Path.GetExtension(this.saveExcelPath);
            File.Copy(this.tplPath, fileName, true);

            //set excel object
            Workbook excel = Utils.OpenExcel(asposeLicPath, fileName);
            //Workbook excel = Utils.NewExcel(asposeLicPath);
            Worksheet sheet = excel.Worksheets[this.sheetName];
            Cells cells = sheet.Cells;

            //set variables
            var addRow = 0; //寫入筆數
            //var colLen = dt.Columns.Count;  //匯入excel檔的欄位數
            //var colLen = dt.Columns.Count;  //匯入excel檔的欄位數
            //var colLen2 = colLen + 1;       //error excel檔的欄位數要加1

            //column index to string
            var colStrs = new string[colLen + 1];
            for (var i = 0; i <= colLen; i++)
                colStrs[i] = ColNumToStr(i + 1);

            //寫入error excel
            foreach (var errorRow in _errorRows)
            {
                //copy/paste row
                var rowPos = startRow + addRow;
                if (addRow > 0)
                {
                    //add row & copy row height
                    cells.InsertRow(rowPos - 1);   //base 1 !! insert after
                    cells.SetRowHeight(rowPos - 1, cells.GetRowHeight(startRow)); //base 1 !!

                    //copy styles(background/foreground colors, fonts, alignment styles etc.)
                    for (int col = 0; col <= colLen; col++)
                        cells[rowPos, col].Style = cells[startRow, col].Style;    //base 1 !!
                }

                //fill row
                var dtRow = this.dt.Rows[errorRow.Key];
                var rowStr = rowPos.ToString();
                for (var i = 0; i < colLen; i++)
                    cells[colStrs[i] + rowStr].PutValue(dtRow[i].ToString());

                //write error msg
                cells[colStrs[colLen] + rowStr].PutValue(errorRow.Value);
                addRow++;
            }
            excel.Save(fileName);
        }

        /// <summary>
        /// 檢查目前這一筆資料的狀態, true(正確), false(有誤)
        /// </summary>
        /// <returns></returns>
        protected bool NowRowStatus()
        {
            return (this.rowError.Count() == 0);
        }

        public void AddOk(int rowNo)
        {
            _okRowNos.Add(rowNo);
        }

        //excel column number to letter string
        //colNo: base 1
        private string ColNumToStr(int colNo)
        {
            int div = colNo;
            string colStr = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colStr = (char)(65 + mod) + colStr;
                div = (int)((div - mod) / 26);
            }
            return colStr;
        }

        #region 欄位值 parsing function
        protected int? ParseIntAndLog(string colName, string value)
        {
            if (String.IsNullOrEmpty(value))
                return (int?)null;

            int value2;
            if (int.TryParse(value, out value2))
                return value2;

            rowError.Add(colName + _inputError);
            return (int?)null;
        }

        protected decimal? ParseDecimalAndLog(string colName, string value)
        {
            //return String.IsNullOrEmpty(value) ? (decimal?)null : decimal.Parse(value);
            if (String.IsNullOrEmpty(value))
                return (decimal?)null;

            decimal value2;
            if (decimal.TryParse(value, out value2))
                return value2;

            rowError.Add(colName + _inputError);
            return (decimal?)null;
        }

        protected DateTime? ParseDatetimeAndLog(string colName, string value)
        {
            if (String.IsNullOrEmpty(value))
                return (DateTime?)null;

            DateTime value2;
            if (DateTime.TryParse(value, out value2))
                return value2;

            rowError.Add(colName + _inputError);
            return (DateTime?)null;
        }
        #endregion
        */

        #endregion

    }//class
}
