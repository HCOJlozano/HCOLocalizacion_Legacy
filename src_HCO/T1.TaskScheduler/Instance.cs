using System;
using Quartz;
using Quartz.Impl;
using SAPbobsCOM;
using T1.B1;
using T1.B1.AsignacionTercerosAsientos;

namespace T1.TaskScheduler
{
    public class Instance
    {
        static private Instance _TaskScheduler = new Instance();
        static private IScheduler _scheduler = null;

        private Instance()
        {
            try
            {
                _scheduler = StdSchedulerFactory.GetDefaultScheduler();
                _scheduler.Start();

                RegisterAllJobs();
            }
            catch (Exception er)
            {
               
            }
        }

        public static void RegisterAllJobs()
        {
            try
            {
                var _Job = JobBuilder.Create<B1.AsignacionTercerosAsientos.Main>().WithIdentity("T1.B1.AsigTerc.J01", "T1.B1.AsigTerc.G01").Build();
                var _Trigger = TriggerBuilder.Create().WithIdentity("T1.B1.AsigTerc.T01", "T1.B1.AsigTerc.G01").StartNow().WithCronSchedule("0 0/2 * * * ?").Build();

                var _JobAsset = JobBuilder.Create<B1.AsignacionActivosAsientos.Main>().WithIdentity("T1.B1.AsigAsset.J01", "T1.B1.AsigAsset.G01").Build();
                var _TriggerAsset = TriggerBuilder.Create().WithIdentity("T1.B1.AsigAsset.T01", "T1.B1.AsigAsset.G01").StartNow().WithCronSchedule("0 0/20 * * * ?").Build();

                if (!isJobRegistered(_Job.Key))
                    addJob(_Job, _Trigger, "10");

                if (!isJobRegistered(_JobAsset.Key))
                    addJob(_JobAsset, _TriggerAsset, "20");
            }
            catch (Exception er)
            {

            }
        }

        public static Instance TaskScheduler()
        {
            return _TaskScheduler;
        }

        static public void StopCron()
        {
            _scheduler.Shutdown();
        }

        static public void addJob(IJobDetail Job, ITrigger Trigger, string time)
        {
            _scheduler.ScheduleJob(Job, Trigger);
            addReferenceJobToDB(Trigger.Key.Name, time);
        }

        static public void addReferenceJobToDB(string name, string taskTime)
        {
            try
            {
                var queryCheckRecord = string.Format(Queries.Instance.Queries().Get("CheckTaskScheduler"), name);
                var recordSet = (Recordset)MainObject.Instance.B1Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordSet.DoQuery(queryCheckRecord);

                if (recordSet.RecordCount == 0)
                {
                    var queryInsertRecord = string.Format(Queries.Instance.Queries().Get("InsertInitialRecordTaskScheduler"), name, taskTime);
                    recordSet.DoQuery(queryInsertRecord);
                }
            }
            catch(Exception ex)
            {

            }
        }

        static public void pauseTrigger(string TriggerName, string TriggerGroup)
        {

            TriggerKey objTrigerKey = new TriggerKey(TriggerName, TriggerGroup);
            TriggerState objState = _scheduler.GetTriggerState(objTrigerKey);
            _scheduler.PauseTrigger(objTrigerKey);

        }

        static public void continueTrigger(string TriggerName, string TriggerGroup)
        {

            TriggerKey objTrigerKey = new TriggerKey(TriggerName, TriggerGroup);
            TriggerState objState = _scheduler.GetTriggerState(objTrigerKey);
            _scheduler.ResumeTrigger(objTrigerKey);

        }

        static public TriggerState getTriggerStatus(string TriggerName, string TriggerGroup)
        {

            TriggerKey objTrigerKey = new TriggerKey(TriggerName, TriggerGroup);
            return _scheduler.GetTriggerState(objTrigerKey);
        }

        static public bool isJobRegistered(JobKey objJobKey)
        {
            return _scheduler.CheckExists(objJobKey);
        }
    }
}
