using System.Collections.Generic;
using System.Globalization;
using System;
using System.ComponentModel;
using System.Linq;
using Base.Services;

using System.Data;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using Base;

namespace SoftNetWebII.Models
{
    public class API_EStore_Open
    {

        public int x { get; set; } 
        public int y { get; set; }
        public int cabinetNo { get; set; } 
        public API_EStore_Open(int cabinetNo,int x, int y)
        {
            this.x = x;
            this.y = y;
            this.cabinetNo = cabinetNo;
        }
    }
    public class API_EStore_Close
    {
        public int x { get; set; } = -1;
        public int y { get; set; } = -1;
        public string cabinetNo { get; set; } = "";
        public int weight { get; set; } = 0;
    }
    public class API_EStore_CallBack
    {
        public bool isFinished { get; set; }
        public List<API_EStore_Close> result { get; set; }
    }

    public class GetPI01
    {
        public string lastOpreateTime { get; set; }
        public string mac { get; set; }
        public int power { get; set; }
        public string routerid { get; set; }
        public int rssi { get; set; }
        public string showStyle { get; set; }
    }

    public class SetPI01
    {
        public string mac;
        public string ledrgb;
        public string 料号;
    }

    public class API_EnterKeyResult
    {
        public string mac { get; set; }
        public int result { get; set; }
    }
    public class API_UpdateTagResult
    {
        public string cmdtoken { get; set; }
        public int lednum { get; set; }
        public string mac { get; set; }
        public string message { get; set; }
        public bool result { get; set; }
    }
    public class API_SetTagType3
    {
        public string mac { get; set; }
        public string reserve { get; set; }
        public int ledmode { get; set; }
        public int buzzer { get; set; }
        public string ledrgb { get; set; }
        public string lednum { get; set; }
        public int outtime { get; set; }
    }


    public class DiaryEvent
    {

        public string id;
        /**名稱*/
        public string name;
        /**內容*/
        public string content;
        /**連結*/
        public string url;
        /**圖片連結*/
        public string imgUrl;
        /**開始時間*/
        public string startDate;
        /**結束時間*/
        public string endDate;

        public string textColor;
        public DiaryEvent(string id, string name,string content, string startDate, string endTMP, string textColor)
        {
            this.id= id;
            this.name = name;
            this.content = content;
            this.startDate = startDate;
            this.endDate = endTMP;
            this.textColor = textColor;
        }

        public static List<DiaryEvent> GetDateRange(DateTime start, DateTime end)
        {
            List<DiaryEvent> result = new List<DiaryEvent>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData($"select * from SoftNetSYSDB.[dbo].[APS_WorkTimeNote] where CalendarDate>='{start.ToString("yyyy/MM/dd HH:mm:ss")}' and CalendarDate<='{end.ToString("yyyy/MM/dd HH:mm:ss")}' order by NeedId,SimulationId,StationNO,CalendarDate");
                if (dt != null && dt.Rows.Count > 0)
                {
                    bool run = true;
                    DateTime endTMP01 = new DateTime();
                    DateTime endTMP03 = new DateTime();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DataRow dr = dt.Rows[i];
                        if ((i + 1) < dt.Rows.Count && dt.Rows[(i + 1)]["SimulationId"].ToString() == dr["SimulationId"].ToString())
                        {
                            if (run)
                            {
                                endTMP01 = Convert.ToDateTime(dr["CalendarDate"]);
                                run = false;
                            }
                            continue;
                        }
                        else
                        {
                            if (run)
                            {
                                endTMP01 = Convert.ToDateTime(dr["CalendarDate"]);
                            }
                            if (bool.Parse(dr["Type1"].ToString())) { endTMP03 = Convert.ToDateTime(dr["CalendarDate"]).AddSeconds(int.Parse(dr["Time1_C"].ToString())); }
                            if (bool.Parse(dr["Type2"].ToString())) { endTMP03 = Convert.ToDateTime(dr["CalendarDate"]).AddSeconds(int.Parse(dr["Time2_C"].ToString())); }
                            if (bool.Parse(dr["Type3"].ToString())) { endTMP03 = Convert.ToDateTime(dr["CalendarDate"]).AddSeconds(int.Parse(dr["Time3_C"].ToString())); }
                            if (bool.Parse(dr["Type4"].ToString())) { endTMP03 = Convert.ToDateTime(dr["CalendarDate"]).AddSeconds(int.Parse(dr["Time4_C"].ToString())); }
                            result.Add(new DiaryEvent(dr["SimulationId"].ToString(), $"{dr["StationNO"].ToString()} {dr["NeedId"].ToString()}", dr["DOCNumberNO"].ToString().Trim(), endTMP01.ToString("s"), endTMP03.ToString("s"), "pink"));
                            run = true;
                        }
                    }
                }
            }
            return result;
        }
        public static List<DiaryEvent> LoadAllAppointmentsInDateRange(double start, double end)
        {
            var fromDate = ConvertFromUnixTimestamp(start);
            var toDate = ConvertFromUnixTimestamp(end);
            List<DiaryEvent> result = new List<DiaryEvent>();
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                /*
                DataTable dt = db.DB_GetData($"select * from FullCalendarMVC_Demo.[dbo].[AppointmentDiary] where DateTimeScheduled>='{fromDate.ToString("yyyy/MM/dd HH:mm:ss")}' and DATEADD(minute,AppointmentLength,DateTimeScheduled)<='{toDate.ToString("yyyy/MM/dd HH:mm:ss")}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        DiaryEvent rec = new DiaryEvent();
                        rec.ID = int.Parse(dr["ID"].ToString());
                        rec.SomeImportantKeyID = int.Parse(dr["SomeImportantKey"].ToString());
                        rec.StartDateString = Convert.ToDateTime(dr["DateTimeScheduled"]).ToString("s"); // "s" is a preset format that outputs as: "2009-02-27T12:12:22"
                        rec.EndDateString = Convert.ToDateTime(dr["DateTimeScheduled"].ToString()).AddMinutes(int.Parse(dr["AppointmentLength"].ToString())).ToString("s"); // field AppointmentLength is in minutes
                        rec.Title = dr["Title"].ToString() + " - " + dr["AppointmentLength"].ToString() + " mins";
                        rec.StatusString = Enums.GetName<AppointmentStatus>((AppointmentStatus)int.Parse(dr["StatusENUM"].ToString()));
                        rec.StatusColor = Enums.GetEnumDescription<AppointmentStatus>(rec.StatusString);
                        string ColorCode = rec.StatusColor.Substring(0, rec.StatusColor.IndexOf(":"));
                        rec.ClassName = rec.StatusColor.Substring(rec.StatusColor.IndexOf(":") + 1, rec.StatusColor.Length - ColorCode.Length - 1);
                        rec.StatusColor = ColorCode;
                        result.Add(rec);
                    }
                }
                */
            }
            return result;

        }


        public static List<DiaryEvent> LoadAppointmentSummaryInDateRange(double start, double end)
        {

            var fromDate = ConvertFromUnixTimestamp(start);
            var toDate = ConvertFromUnixTimestamp(end);
            List<DiaryEvent> result = new List<DiaryEvent>();
            /*
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                string _s = $"select * from FullCalendarMVC_Demo.[dbo].[AppointmentDiary] where DateTimeScheduled>='{fromDate.ToString("yyyy/MM/dd HH:mm:ss")}' and DATEADD(minute,AppointmentLength,DateTimeScheduled)<='{toDate.ToString("yyyy/MM/dd HH:mm:ss")}'";
                DataTable dt = db.DB_GetData($"select * from FullCalendarMVC_Demo.[dbo].[AppointmentDiary] where DateTimeScheduled>='{fromDate.ToString("yyyy/MM/dd HH:mm:ss")}' and DATEADD(minute,AppointmentLength,DateTimeScheduled)<='{toDate.ToString("yyyy/MM/dd HH:mm:ss")}'");
                if (dt != null && dt.Rows.Count > 0)
                {
                    int i = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        DiaryEvent rec = new DiaryEvent();
                        rec.ID = i;
                        rec.SomeImportantKeyID = -1;

                        string StringDate = string.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(dr["DateTimeScheduled"].ToString()));
                        rec.StartDateString =$"{StringDate}T00:00:00"; //ISO 8601 format
                        rec.EndDateString = $"{StringDate}T23:59:59";
                        rec.Title = "Booked: " + i.ToString();
                        result.Add(rec);
                        i++;
                    }
                }
            }
            
            */
            /*
            DiaryEvent rec = new DiaryEvent();
            rec.ID = 1;
            rec.SomeImportantKeyID = -1;
            rec.StartDateString = "2022-09-08T00:00:00";
            rec.EndDateString = "2022-09-08T23:59:59";
            rec.Title = "Booked: 1";
            result.Add(rec);
            result[0].StartDateString = "2022-09-08T00:00:00";
            */
            return result;
            /*
            using (DiaryContainer ent = new DiaryContainer())
            {
                var rslt = ent.AppointmentDiary.Where(s => s.DateTimeScheduled >= fromDate && System.Data.Objects.EntityFunctions.AddMinutes(s.DateTimeScheduled, s.AppointmentLength) <= toDate)
                                                        .GroupBy(s => System.Data.Objects.EntityFunctions.TruncateTime(s.DateTimeScheduled))
                                                        .Select(x => new { DateTimeScheduled = x.Key, Count = x.Count() });

                List<DiaryEvent> result = new List<DiaryEvent>();
                int i = 0;
                foreach (var item in rslt)
                {
                    DiaryEvent rec = new DiaryEvent();
                    rec.ID = i; //we dont link this back to anything as its a group summary but the fullcalendar needs unique IDs for each event item (unless its a repeating event)
                    rec.SomeImportantKeyID = -1;
                    string StringDate = string.Format("{0:yyyy-MM-dd}", item.DateTimeScheduled);
                    rec.StartDateString = StringDate + "T00:00:00"; //ISO 8601 format
                    rec.EndDateString = StringDate + "T23:59:59";
                    rec.Title = "Booked: " + item.Count.ToString();
                    result.Add(rec);
                    i++;
                }

                return result;
            }
            */
            return null;
        }

        public static void UpdateDiaryEvent(int id, string NewEventStart, string NewEventEnd)
        {
            // EventStart comes ISO 8601 format, eg:  "2000-01-10T10:00:00Z" - need to convert to DateTime
            /*
            using (DiaryContainer ent = new DiaryContainer())
            {
                var rec = ent.AppointmentDiary.FirstOrDefault(s => s.ID == id);
                if (rec != null)
                {
                    DateTime DateTimeStart = Convert.ToDateTime(NewEventStart, null, DateTimeStyles.RoundtripKind).ToLocalTime(); // and convert offset to localtime
                    rec.DateTimeScheduled = DateTimeStart;
                    if (!String.IsNullOrEmpty(NewEventEnd))
                    {
                        TimeSpan span = Convert.ToDateTime(NewEventEnd, null, DateTimeStyles.RoundtripKind).ToLocalTime() - DateTimeStart;
                        rec.AppointmentLength = Convert.ToInt32(span.TotalMinutes);
                    }
                    ent.SaveChanges();
                }
            }
            */
        }


        private static DateTime ConvertFromUnixTimestamp(double timestamp)
        {
            var origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return origin.AddSeconds(timestamp);
        }


        public static bool CreateNewEvent(string Title, string NewEventDate, string NewEventTime, string NewEventDuration)
        {
            /*
            try
            {
                DiaryContainer ent = new DiaryContainer();
                AppointmentDiary rec = new AppointmentDiary();
                rec.Title = Title;
                rec.DateTimeScheduled = DateTime.ParseExact(NewEventDate + " " + NewEventTime, "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
                rec.AppointmentLength = Int32.Parse(NewEventDuration);
                ent.AppointmentDiary.Add(rec);
                ent.SaveChanges();
            }
            catch (Exception)
            {
                return false;
            }
            return true;
            */
            return false;
        }
    }

    public enum AppointmentStatus
    {
        [Description("#01DF3A:ENQUIRY")] // green
        Enquiry = 0,
        [Description("#FF8000:BOOKED")] // orange
        Booked,
        [Description("#FF0000:CONFIRMED")] // red
        Confirmed

    }

    public static class Enums
    {
        /// Get all values
        public static IEnumerable<T> GetValues<T>()
        {
            return Enum.GetValues(typeof(T)).Cast<T>();
        }

        /// Get all the names
        public static IEnumerable<T> GetNames<T>()
        {
            return Enum.GetNames(typeof(T)).Cast<T>();
        }

        /// Get the name for the enum value
        public static string GetName<T>(T enumValue)
        {
            return Enum.GetName(typeof(T), enumValue);
        }

        /// Get the underlying value for the Enum string
        public static int GetValue<T>(string enumString)
        {
            return (int)Enum.Parse(typeof(T), enumString.Trim());
        }

        public static string GetEnumDescription<T>(string value)
        {
            Type type = typeof(T);
            var name = Enum.GetNames(type).Where(f => f.Equals(value, StringComparison.CurrentCultureIgnoreCase)).Select(d => d).FirstOrDefault();

            if (name == null)
            {
                return string.Empty;
            }
            var field = type.GetField(name);
            var customAttribute = field.GetCustomAttributes(typeof(DescriptionAttribute), false);
            return customAttribute.Length > 0 ? ((DescriptionAttribute)customAttribute[0]).Description : name;
        }
    }
}
