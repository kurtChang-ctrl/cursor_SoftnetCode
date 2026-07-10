using Base;
using Base.Models;
using Base.Services;
using BaseApi.Controllers;
using BaseApi.Services;
using Microsoft.AspNetCore.Mvc;
using SoftNetWebII.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace SoftNetWebII.Controllers
{
    public class WarRoomController : ApiCtrl
    {
        private SFC_Common _SFC_Common = null;

        public WarRoomController(SFC_Common sfc_Common)
        {
            if (_SFC_Common == null)
            {
                _SFC_Common = sfc_Common;
            }
        }

        /// <summary>
        /// 功能1：製造現場工作人員戰情室
        /// </summary>
        public ActionResult Manufacturing()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            return View();
        }

        /// <summary>
        /// 功能2：生管人員戰情室
        /// </summary>
        public ActionResult ProductionManagement()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            return View();
        }

        /// <summary>
        /// 功能3：高階管理戰情室
        /// </summary>
        public ActionResult Executive()
        {
            var br = _Fun.GetBaseUser();
            if (br == null || !br.IsLogin || br.UserNO == "")
            {
                return RedirectToAction("Login", "Home", new { url = _Http.GetWebUrl() });
            }
            return View();
        }

        private WarRoomService Service()
        {
            return new WarRoomService(Ctrl, _SFC_Common);
        }

        #region API Methods

        /// <summary>
        /// 取得今日/昨日產量統計（區分半成品與成品）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetProductionOutput(string dateType = "today")
        {
            return JsonToCnt(await Service().GetProductionOutputAsync(dateType));
        }

        /// <summary>
        /// 取得工單統計（進行中、已完成、達成率）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetWorkOrderStats()
        {
            return JsonToCnt(await Service().GetWorkOrderStatsAsync());
        }

        /// <summary>
        /// 取得工站即時狀態（支援多工單）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetStationStatus()
        {
            return JsonToCnt(await Service().GetStationStatusAsync());
        }

        /// <summary>
        /// 取得即時警示（主動干涉引擎）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetActiveInterventionAlerts(int limit = 10)
        {
            return JsonToCnt(await Service().GetActiveInterventionAlertsAsync(limit));
        }

        /// <summary>
        /// 取得我的工單任務
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetMyWorkOrders(string stationNO)
        {
            return JsonToCnt(await Service().GetMyWorkOrdersAsync(stationNO));
        }

        /// <summary>
        /// 取得物料需求狀態
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetMaterialRequirements(string orderNO)
        {
            return JsonToCnt(await Service().GetMaterialRequirementsAsync(orderNO));
        }

        /// <summary>
        /// 取得所有工站列表
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetAllStations()
        {
            return JsonToCnt(await Service().GetAllStationsAsync());
        }

        /// <summary>
        /// 取得生產現況總覽（功能2）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetProductionOverview()
        {
            return JsonToCnt(await Service().GetProductionOverviewAsync());
        }

        /// <summary>
        /// 取得未來7天排程預警（功能2）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetScheduleWarnings(int days = 7)
        {
            return JsonToCnt(await Service().GetScheduleWarningsAsync(days));
        }

        /// <summary>
        /// 取得工站負荷分析（功能2）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetStationLoadAnalysis(string dateType = "today")
        {
            return JsonToCnt(await Service().GetStationLoadAnalysisAsync(dateType));
        }

        /// <summary>
        /// 取得物料庫存預警（功能2）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetMaterialStockWarnings()
        {
            return JsonToCnt(await Service().GetMaterialStockWarningsAsync());
        }

        /// <summary>
        /// 取得工單執行狀態（功能2）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetWorkOrderStatus(string status = "all")
        {
            return JsonToCnt(await Service().GetWorkOrderStatusAsync(status));
        }

        /// <summary>
        /// 取得主動干涉事件統計（功能2）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetInterventionEventStats(string dateType = "today")
        {
            return JsonToCnt(await Service().GetInterventionEventStatsAsync(dateType));
        }

        /// <summary>
        /// 取得核心KPI總覽（功能3）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetCoreKPIs(string period = "month")
        {
            return JsonToCnt(await Service().GetCoreKPIsAsync(period));
        }

        /// <summary>
        /// 取得生產效能趨勢分析（功能3）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetProductionTrend(string granularity = "month", int months = 6)
        {
            return JsonToCnt(await Service().GetProductionTrendAsync(granularity, months));
        }

        /// <summary>
        /// 取得產能利用率趨勢（功能3）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetCapacityUtilizationTrend(string granularity = "month", int months = 6)
        {
            return JsonToCnt(await Service().GetCapacityUtilizationTrendAsync(granularity, months));
        }

        /// <summary>
        /// 取得進步/退步分析（功能3）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetProgressAnalysis(string period = "month")
        {
            return JsonToCnt(await Service().GetProgressAnalysisAsync(period));
        }

        /// <summary>
        /// 取得成本效益分析（功能3）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetCostBenefitAnalysis(string period = "month")
        {
            return JsonToCnt(await Service().GetCostBenefitAnalysisAsync(period));
        }

        /// <summary>
        /// 取得目標達成狀況（功能3）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetTargetAchievement(string period = "month")
        {
            return JsonToCnt(await Service().GetTargetAchievementAsync(period));
        }

        /// <summary>
        /// 取得主動干涉次數（功能3）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetInterventionCount(int days = 30)
        {
            return JsonToCnt(await Service().GetInterventionCountAsync(days));
        }

        /// <summary>
        /// 取得主動干涉明細（功能3）
        /// </summary>
        [HttpPost]
        public async Task<ContentResult> GetInterventionDetail(int days = 30)
        {
            return JsonToCnt(await Service().GetInterventionDetailAsync(days));
        }

        #endregion
    }
}
