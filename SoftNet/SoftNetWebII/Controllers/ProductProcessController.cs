using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using BaseWeb.Models;
using BaseWeb.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

using SoftNetWebII.Services;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class ProductProcessController : ApiCtrl
    {
        public async Task<ActionResult> Read()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            //test
            //_Fun.Except("exception test");
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                List<IdStrDto> apply_PP_Name = new List<IdStrDto>();
                apply_PP_Name.Add(new IdStrDto("", ""));

                DataTable dt = db.DB_GetData($"SELECT PP_Name FROM SoftNetSYSDB.[dbo].[PP_ProductProcess] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        apply_PP_Name.Add(new IdStrDto(dr["PP_Name"].ToString(), dr["PP_Name"].ToString()));
                    }
                }
                ViewBag.Apply_PP_Name = apply_PP_Name;
                List<IdStrDto> factory = new List<IdStrDto>();
                dt = db.DB_GetData($"SELECT FactoryName FROM SoftNetMainDB.[dbo].[Factory] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        factory.Add(new IdStrDto(dr["FactoryName"].ToString(), dr["FactoryName"].ToString()));
                    }
                }
                ViewBag.FactoryName = factory;
                List<IdStrDto> stationNO = new List<IdStrDto>();
                stationNO.Add(new IdStrDto("", ""));
                dt = db.DB_GetData($"SELECT StationNO FROM SoftNetSYSDB.[dbo].[PP_Station] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        stationNO.Add(new IdStrDto(dr["StationNO"].ToString(), dr["StationNO"].ToString()));
                    }
                }
                ViewBag.StationNO = stationNO;
                List<IdStrDto> mFNO = new List<IdStrDto>();
                mFNO.Add(new IdStrDto("", ""));
                dt = db.DB_GetData($"SELECT MFNO FROM SoftNetMainDB.[dbo].[MFData] where ServerId='{_Fun.Config.ServerId}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        mFNO.Add(new IdStrDto(dr["MFNO"].ToString(), dr["MFNO"].ToString()));
                    }
                }
                ViewBag.MFNO = mFNO;
                List<IdStrDto> calendarName = new List<IdStrDto>();
                dt = db.DB_GetData($"SELECT CalendarName FROM SoftNetSYSDB.[dbo].[PP_HolidayCalendar] where ServerId='{_Fun.Config.ServerId}' group by CalendarName");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        calendarName.Add(new IdStrDto(dr["CalendarName"].ToString(), dr["CalendarName"].ToString()));
                    }
                }
                ViewBag.CalendarName = calendarName;


                List<IdStrDto> isYesNo = new List<IdStrDto>();
                isYesNo.Add(new IdStrDto("0", "否"));
                isYesNo.Add(new IdStrDto("1", "是"));
                ViewBag.IsYesNo = isYesNo;
            }
            return View();
        }

        [HttpPost]
        public async Task<JsonResult> ImportExcel(IFormFile file)
        {

            ResultImportDto dd = new ResultImportDto();
            dd.ErrorMsg = "sssss";
            if (dd != null) { Json(dd); }
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                /*
                db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[PP_ProductProcess] where PP_Name like '%_加工製程'");
                db.DB_SetData($"delete from SoftNetSYSDB.[dbo].[PP_ProductProcess_Item] where PP_Name like '%_加工製程'");
                DataTable dt = db.DB_GetData("SELECT *  FROM SoftNetMainDB.[dbo].[BOM] where Apply_PP_Name like '%_加工製程'");
                if (dt!=null && dt.Rows.Count > 0)
                {
                    foreach(DataRow dr in dt.Rows)
                    {
                        db.DB_SetData($"delete from SoftNetMainDB.[dbo].[BOMII] where BOMId='{dr["Id"].ToString()}'");
                    }
                }
                db.DB_SetData($"delete FROM SoftNetMainDB.[dbo].[BOM] where Apply_PP_Name like '%_加工製程'");
                */

                /*
                DataTable dt = db.DB_GetData("SELECT *  FROM [SoftNetMainDB].[dbo].[BOMII] where PartNO like '%-N'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    string tmp = "";
                    foreach (DataRow dr in dt.Rows)
                    {
                        tmp = dr["PartNO"].ToString().Replace("-N","");
                        db.DB_SetData($"delete from [SoftNetMainDB].[dbo].[BOMII] where PartNO like '{tmp}%'");
                    }
                }
                */

            }
            
            string uname = "";
            if (_Fun.GetBaseUser() != null)
            { uname = _Fun.GetBaseUser().UserName; }
            var importDto = new ExcelImportDto<SimulationDto>()
            {
                //###??? 依需求命名   ImportType = "PPName_AND_BOM",
                ImportType = "萬_PPName_AND_BOM_二次",
                TplPath = _Xp.DirTpl + "UserImport.xlsx",
                FnSaveImportRows = SaveImportRows,
                CreatorName = uname,
            };
            string dirUpload = @"D:\書_sampleCode\HrAdm-master\_upload\UserImport\";
            var model = await _WebExcel.ImportByFileAsync(file, dirUpload, importDto);
            return Json(model);
        }
        private List<string> SaveImportRows(List<SimulationDto> okRows)
        {
            var results = new List<string>();


            /*
            var db = _Xp.GetDb();
            var deptIds = db.Dept.Select(a => a.DeptId).ToList();
            foreach (var row in okRows)
            {
                //check rules: deptId
                if (!deptIds.Contains(row.DeptId))
                {
                    results.Add("DeptId Wrong.");
                    continue;
                }

                //check rules: Account not repeat
                if (db.User.Any(a => a.Account == row.Account))
                {
                    results.Add("Account Existed.");
                    continue;
                }

                #region set entity model & save db
                db.User.Add(new User()
                {
                    Id = _Str.NewId('Z'),
                    Name = row.Name,
                    Account = row.Account,
                    Pwd = row.Pwd,
                    DeptId = row.DeptId,
                    Status = true,
                });

                //save db
                try
                {
                    db.SaveChanges();
                    results.Add("");
                }
                catch (Exception ex)
                {
                    results.Add(ex.InnerException.Message);
                }
                #endregion
            }

            db.Dispose();
            */
            return results;
        }


        [HttpPost]
        public async Task<ContentResult> GetPage(DtDto dt)
        {
            return JsonToCnt(await EditService().GetPageAsync(Ctrl, dt));
        }

        private ProductProcessService EditService()
        {
            return new ProductProcessService(Ctrl);
        }

        //讀取要修改的資料(Get Updated Json)
        [HttpPost]
        public async Task<ContentResult> GetUpdJson(string key)
        {
            return JsonToCnt(await EditService().GetUpdJsonAsync(key));
        }

        //新增(DB)
        public async Task<JsonResult> Create(string json)
        {
            return Json(await EditService().CreateAsync(_Str.ToJson(json)));
        }
        //修改(DB)
        public async Task<JsonResult> Update(string key, string json)
        {
            return Json(await EditService().UpdateAsync(key, _Str.ToJson(json)));
        }

        //刪除(DB)
        public async Task<JsonResult> Delete(string key)
        {
            return Json(await EditService().DeleteAsync(key));
        }
    }
}
