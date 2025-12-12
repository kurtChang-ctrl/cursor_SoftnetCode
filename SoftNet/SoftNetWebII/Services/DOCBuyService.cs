using Base;
using Base.Enums;
using Base.Models;
using Base.Services;
using BaseApi.Services;
using BaseWeb.Services;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;

using SoftNetWebII.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace SoftNetWebII.Services
{
    public class DOCBuyService : XgEdit
    {
        public DOCBuyService(string ctrl) : base(ctrl) { }
 
        override public EditDto GetDto()
        {
            return new EditDto
            {
                Table = "dbo.[DOC1Buy]",
                PkeyFid = "DOCNumberNO",
                PkeyFids = new string[] { "ServerId", "DOCNumberNO" },
                Col4 = null,
                HasFlowSign = true,
                Items = new EitemDto[] {
                    new() { Fid = "DOCNO" },
                    new() { Fid = "ServerId" },
                    new() { Fid = "DOCNumberNO" },
                    new() { Fid = "DOCDate", Required = true  },
                    new() { Fid = "SourceNO" },
                    new() { Fid = "UserId" },
                    new() { Fid = "SourceNO" },
                    new() { Fid = "MFNO" },
                    new() { Fid = "FlowLevel", Value = "1" },
                    new() { Fid = "FlowStatus", Value = "0" },
                },
                Childs = new EditDto[]
                {
                    new()
                    {
                        IsKeyValueNUM = true,
                        ReadSql = "SELECT A.Id,A.DOCNumberNO,A.PartNO,B.PartName,B.Specification,A.Price,A.Unit,A.QTY,A.SimulationId,A.ArrivalDate,A.Remark FROM SoftNetMainDB.[dbo].[DOC1BuyII] as A,dbo.Material as B where A.DOCNumberNO={0} and A.PartNO=B.PartNO order by A.Id,A.DOCNumberNO",
                        Table = "dbo.[DOC1BuyII]",
                        PkeyFid = "Id",
                        FkeyFid = "DOCNumberNO",
                        OrderBy = "DOCNumberNO",
                        Col4 = null,
                        Items = new EitemDto[] {
                            new() { Fid = "Id" },
                            new() { Fid = "DOCNumberNO" },
                            new() { Fid = "PartNO", Required = true  },
                            new() { Fid = "Price" },
                            new() { Fid = "Unit" },
                            new() { Fid = "QTY" },
                            new() { Fid = "ArrivalDate" },
                            new() { Fid = "SimulationId" },
                            new() { Fid = "Remark" },
                        },
                    },
                },
            };
        }


        private readonly ReadDto dto = new()
        {
            ReadSql = $@"SELECT A.DOCNumberNO,D.Name_zhTW as SignStatusName,A.DOCNO,C.DOCName,A.DOCDate,A.SourceNO,A.UserId,A.DOCType,A.SourceNO,A.MFNO,B.MFName,A.TotalMoney,A.TaxMoney,(A.MFNO+' '+B.MFName) as MFNOName
                        FROM SoftNetMainDB.[dbo].[DOC1Buy] as A
                        Join SoftNetMainDB.[dbo].[MFData] as B On A.MFNO = B.MFNO
                        Join SoftNetMainDB.[dbo].DOCRole as C On A.DOCNO=C.DOCNO
                        join SoftNetMainDB.[dbo].XpCode as D  on D.Type='FlowStatus' and A.FlowStatus=D.Value
                        where A.ServerId='{_Fun.Config.ServerId}' order by A.DOCNO,A.DOCNumberNO",
            Items = new QitemDto[] {
                    new() { Fid = "DOCNO", Op = ItemOpEstr.Like },
                    new() { Fid = "DOCDate" },
                    new() { Fid = "MFNO", Op = ItemOpEstr.Like },
            },
        };

        public async Task<JObject> GetPageAsync(string ctrl, DtDto dt)
        {
            return await new CrudRead().GetPageAsync(dto, dt, ctrl);
        }

        private JObject _inputRow;
        private async Task<string> FnCreateSignRowsAsync(Db db, JObject newKeyJson)
        {
            var newKey = _Str.ReadNewKeyJson(newKeyJson);
            return await _XgFlow.CreateSignRowsAsync(_inputRow, "UserId", "DOCSet_人工填單", newKey, false, db);
        }
        public async Task<ResultDto> OtherCrateAsync(JObject json)
        {
            var rows = json["_rows"] as JArray;
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataRow dr = db.DB_GetFirstDataByDataRow($"SELECT DOCNumberNO FROM SoftNetMainDB.[dbo].[DOC1Buy] where ServerId='{_Fun.Config.ServerId}' and DOCNO='{rows[0]["DOCNO"].ToString()}' and DOCDate='{Convert.ToDateTime(rows[0]["DOCDate"]).ToString("yyyy-MM-dd")}' order by DOCNumberNO desc");
                if (dr == null)
                {
                    rows[0]["DOCNumberNO"] = $"{rows[0]["DOCNO"].ToString()}{Convert.ToDateTime(rows[0]["DOCDate"]).ToString("yyyyMMdd")}0001";
                }
                else
                {
                    int tmp = int.Parse(dr["DOCNumberNO"].ToString().Trim().Substring(12)) + 1;
                    string dd = tmp.ToString().PadLeft(4, '0');
                    rows[0]["DOCNumberNO"] = $"{rows[0]["DOCNO"].ToString()}{Convert.ToDateTime(rows[0]["DOCDate"]).ToString("yyyyMMdd")}{tmp.ToString().PadLeft(4, '0')}";
                }
            }
            _inputRow = _Json.ReadInputJson0(json);
            var service = EditService();
            var result = await service.CreateAsync(json, null, FnCreateSignRowsAsync);
            if (_Valid.ResultStatus(result))
                await _WebFile.SaveCrudFileAsnyc(json, service.GetNewKeyJson(), _Xp.DirLeave, null, null);

            //ResultDto re = await CreateAsync(json);
            return result;
        }
        public List<SignRowDto> GetSignRows(string id)
        {
            /*
            List<SignRowDto> re = new List<SignRowDto>();
            SignRowDto tmp = new SignRowDto();
            tmp.NodeName = "A";
            tmp.SignerName = "B";
            tmp.SignStatusName = "C";
            tmp.SignTime = "";
            tmp.Note = "D";
            re.Add(tmp);
            return re;
            */

            List<SignRowDto> re = new List<SignRowDto>();
            string sql = $@"select s.NodeName,s.SignerName,s.SignTime,s.Note,c.Name_zhTW from SoftNetMainDB.[dbo].XpFlowSign as s
                            join SoftNetMainDB.[dbo].XpFlow as f on s.FlowId=f.Id
                            join SoftNetMainDB.[dbo].XpCode as c on c.Type='SignStatus' and c.Value=s.SignStatus
                            where f.Code = 'DOCSet_人工填單' and s.SourceId = '{id}'";
            using (DBADO db = new DBADO("1", _Fun.Config.Db))
            {
                DataTable dt = db.DB_GetData(sql);
                if (dt != null)
                {
                    foreach (DataRow d in dt.Rows)
                    {
                        SignRowDto tmp = new SignRowDto();
                        tmp.NodeName = d["NodeName"].ToString();
                        tmp.SignerName = d["SignerName"].ToString();
                        tmp.SignStatusName = d["Name_zhTW"].ToString();
                        if (d.IsNull("SignTime")) { tmp.SignTime = ""; } else { tmp.SignTime = Convert.ToDateTime(d["SignTime"]).ToString(_Fun.CsDtFmt); }
                        tmp.SignTime = d["SignTime"].ToString();
                        tmp.Note = d["Note"].ToString();
                        re.Add(tmp);
                    }
                }
            }
            return re;

            /*
                        var locale = _Xp.GetLocale0();
                        var db = _Xp.GetDb();
                        return (from s in db.XpFlowSign 
                        join f in db.XpFlow on s.FlowId equals f.Id
                        join c in db.XpCode on new { Type = "SignStatus", Value = s.SignStatus }
                            equals new { c.Type, c.Value }
                        where (f.Code == "DOCSet_人工填單" && s.SourceId == id)
                        orderby s.FlowLevel
                        select new SignRowDto()
                        {
                            NodeName = s.NodeName,
                            SignerName = s.SignerName,
                            SignStatusName = _XpCode.GetValue(c, locale),
                            SignTime = (s.SignTime == null)
                                ? "" : s.SignTime.Value.ToString(_Fun.CsDtFmt),
                            Note = s.Note,
                        })
                        .ToList();
                       */
        }
    } //class
}

