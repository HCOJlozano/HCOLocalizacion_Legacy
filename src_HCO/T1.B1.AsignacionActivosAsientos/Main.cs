using Quartz;
using SAPbobsCOM;
using System;


namespace T1.B1.AsignacionActivosAsientos
{
    public class Main : IJob
    {
        private string NameScheduler = "T1.B1.AsigAsset.T01";
        public void Execute(IJobExecutionContext context)
        {
            var query = string.Format(Queries.Instance.Queries().Get("CheckLastHourExecution"), "T1.B1.AsigAsset.T01");
            var oRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery(query);
                oRS.MoveFirst();

            if (oRS.RecordCount > 0)
            {
                if(oRS.Fields.Item("Result").Value.ToString().Equals("0") )
                    UpdateJournalFixedAsset();
            }
        }

        private void UpdateJournalFixedAsset()
        {
            try
            {
                var journal = (JournalEntries)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                var oRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                var query = Queries.Instance.Queries().Get("GetValorizationValue");
                    oRS.DoQuery(Queries.Instance.Queries().Get("GetValorizationValue"));

                if (oRS.RecordCount > 0)
                {
                    while (!oRS.EoF)
                    {
                        if (journal.GetByKey(int.Parse(oRS.Fields.Item("TransId").Value.ToString())))
                        {
                            var area = GetValorizationValue(oRS.Fields.Item("DprArea").Value.ToString());
                            journal.UserFields.Fields.Item("U_HCO_ValAre").Value = area;
                            journal.Update();
                        }

                        oRS.MoveNext();
                    }
                }
                updateTaskInfo(0, "Ok");
            }
            catch (Exception ex)
            {
                updateTaskInfo(-1, ex.Message);
            }
        }

        private void updateTaskInfo(int state, string message)
        {
            var query = string.Format(Queries.Instance.Queries().Get("UpdateRecordTaskScheduler"), NameScheduler, MainObject.Instance.B1Company.UserName, state, message);
            var oRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery(query);
        }

        public string GetValorizationValue(string value)
        {
            switch (value)
            {
                case "01":
                    return "I";
                case "02":
                    return "L";
                default:
                    return "C";
            }
        }
    }
}
