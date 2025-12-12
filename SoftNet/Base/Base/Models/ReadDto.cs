using Base.Services;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;

namespace Base.Models
{
    /// <summary>
    /// for Crud List/Query form
    /// </summary>
    public class ReadDto
    {
        /// <summary>
        /// same as DbReadModel.ColList
        /// </summary>
        public string ColList = "";

        public string Db_forFind_From_AS_Name = "";

        /// <summary>
        /// sql string, column select order must same to client side for datatable sorting !!
        /// </summary>
        public string ReadSql = "";

        /// <summary>
        /// sql string for export excel, default to ReadSql.
        /// </summary>
        public string ExportSql = "";

        /// <summary>
        /// sql use square, as: [from],[where],[group],[order]
        /// (TODO: add [whereCond] for client condition !!)
        /// </summary>
        public bool UseSquare = false;

        /// <summary>
        /// default table alias name
        /// </summary>
        public string TableAs = "";

        /// <summary>
        /// (for AuthType=Data only) user fid, default to _Fun.FindUserFid
        /// </summary>
        public string WhereUserFid = _Fun.WhereUserFid;

        /// <summary>
        /// (for AuthType=Data only) dept fid, default to _Fun.FindDeptFid
        /// </summary>
        public string WhereDeptFid = _Fun.WhereDeptFid;

        /// <summary>
        /// for quick search, include table alias, will get like xx% query
        /// </summary>
        public string[] FindCols;

        /// <summary>
        /// or query for column group, suggest to List more easy !!
        /// </summary>
        public List<List<string>> OrGroups;

        /// <summary>
        /// query condition fields
        /// </summary>
        public QitemDto[] Items;

        //sql 使用預存程序
        public string SQL_StoredProgram = "";

        public ProgramReadDataOBJ SQL_ByProgram = null;

        //當 select 無法明確抓取 from 時, 可指定宣告, 用於Select Count(*) as _count
        public string SpecifySQLFrom = "";

    }//class

    public class ProgramReadDataOBJ : IDisposable
    {
        public int RowCount = 0;
        public JArray rows = null;


        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        private bool disposed = false;
        protected virtual void Dispose(bool disposing)
        {
            if (disposed) { return; }
            if (disposing)
            {
                //清理CLR託管資源
                /*

                 */
            }
            //清理非託管資源,寫在下方,如果有的話
            disposed = true;

        }
    }
}
