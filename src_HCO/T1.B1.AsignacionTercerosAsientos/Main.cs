using Quartz;
using SAPbobsCOM;
using System;

namespace T1.B1.AsignacionTercerosAsientos
{
    public class Main : IJob
    {
        private string NameScheduler = "T1.B1.AsigTerc.T01";

        public void Execute(IJobExecutionContext context)
        {
            var query = string.Format(Queries.Instance.Queries().Get("CheckLastHourExecution"), "T1.B1.AsigTerc.T01");
            var oRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRS.DoQuery(query);
            oRS.MoveFirst();

            if (oRS.RecordCount > 0)
            {
                if (oRS.Fields.Item("Result").Value.ToString().Equals("0"))
                    UpdateJournalThird();
            }
        }

        private void UpdateJournalThird()
        {         
            try
            {
                var journal = (JournalEntries) MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                var oRS = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRS.DoQuery(Queries.Instance.Queries().Get("GetThirdMissingReference"));

                if (oRS.RecordCount > 0)
                {
                    while(!oRS.EoF) 
                    { 
                        if (journal.GetByKey(int.Parse(oRS.Fields.Item("TransId").Value.ToString())))
                        {
                            for(int i=0; i<journal.Lines.Count; i++)
                            {
                                journal.Lines.SetCurrentLine(i);
                                journal.Lines.UserFields.Fields.Item("U_HCO_RELPAR").Value = oRS.Fields.Item("Code").Value;
                            }

                            journal.Update();
                        }

                        oRS.MoveNext();
                    }
                }

                updateTaskInfo(0, "Ok");
            }
            catch(Exception ex)
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
    }
}
