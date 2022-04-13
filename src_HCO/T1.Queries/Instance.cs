using log4net;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace T1.Queries
{
    public class Instance
    {
        static private Instance _Queries = new Instance();
        static private List<Entities.Queries> queriesAll;
        static private List<Entities.Queries> queriesSQL;
        static private List<Entities.Queries> queriesHANA;
        private static readonly string LOG_LEVEL = "Debug";
        private static readonly string PATH_ALL = @"Queries\\ALL.json";
        private static readonly string PATH_HANA = @"Queries\\HANA.json";
        private static readonly string PATH_SQL = @"Queries\\MSSQL.json";
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, LOG_LEVEL);

        private Instance()
        {
            try
            {
                var currentExecutionPath = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                var dataAll = File.ReadAllText(currentExecutionPath+ "\\Queries\\ALL.json");
                var dataHana = File.ReadAllText(currentExecutionPath + "\\Queries\\HANA.json");
                var dataSql = File.ReadAllText(currentExecutionPath + "\\Queries\\MSSQL.json");
                
                if( !dataAll.Equals(string.Empty) )
                    queriesAll = GetQueriesFromJson<Entities.Queries>(dataAll, "Queries");

                if (!dataHana.Equals(string.Empty))
                    queriesHANA = GetQueriesFromJson<Entities.Queries>(dataHana, "Queries");

                if (!dataSql.Equals(string.Empty))
                    queriesSQL = GetQueriesFromJson<Entities.Queries>(dataSql, "Queries");
            }
            catch (Exception er)
            {
                _Logger.Error("COM Error", er);
            }
        }

        public static Instance Queries()
        {
            return _Queries;
        }

        public string Get(string queryName)
        {
            IEnumerable<string> queryToReturn = null;
          
            var asda = T1.B1.MainObject.Instance.B1Company.DbServerType.ToString();
            if (T1.B1.MainObject.Instance.B1Company.DbServerType.ToString().Contains("MSSQL"))
            {
                if (queriesSQL != null)
                    queryToReturn = from q in queriesSQL where q.QueryName == queryName select q.Query;
            }
            else if (T1.B1.MainObject.Instance.B1Company.DbServerType.ToString().Contains("HANA"))
            {
                if (queriesHANA != null)
                    queryToReturn = from q in queriesHANA where q.QueryName == queryName select q.Query;
            }
            
            if(queryToReturn.Count() == 0)
            {
                if (queriesAll != null)
                    queryToReturn = from q in queriesAll where q.QueryName == queryName select q.Query;
            }

            if (queryToReturn.Count() > 0) 
                return queryToReturn.First();

            return string.Empty;
        }

        public List<T> GetQueriesFromJson<T>(string pathToFile, string nodeToRead)
        {
            JObject o = JObject.Parse(pathToFile);
            JArray a = (JArray)o[nodeToRead];
            return a.ToObject<List<T>>();
        }

    }
}
