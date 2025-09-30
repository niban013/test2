using Quartz;
using Quartz.Impl;
using Quartz.Impl.Triggers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DriveAgent.Tool
{
    public class QuartzHelper
    {
        // 調度器工廠 (負責創建和配置 IScheduler 實例)
        private static ISchedulerFactory? factory = null;
        // 調度器 排程 (提供了管理 Job 和 Trigger 以及控制調度器生命週期的方法)
        public static IScheduler? scheduler = null;

        /// <summary>
        /// 創建、啟動排程
        /// </summary>
        public static void CreateScheduler()
        {
            // tdSchedulerFactory 是 ISchedulerFactory 接口的默認實現,用於創建和配置調度器實例
            factory = new StdSchedulerFactory();
            // 獲取了一個 IScheduler 實例 ， 用來管理 Job 和 Trigger
            scheduler = factory?.GetScheduler().Result;
            // 啟動調度器後,可根據註冊的 Trigger 按計劃執行相應的 Job
            scheduler?.Start();
        }

        #region 觸發器1：添加Job
        /// <summary>
        /// 觸發器1：添加Job並以週期的形式運行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="jobName"></param>
        /// <param name="startTime"></param>
        /// <param name="simpleTime"></param>
        /// <param name="jobDataMap"></param>
        /// <returns></returns>
        public static async Task<DateTimeOffset> AddJob<T>(DateTimeOffset startTime, TimeSpan simpleTime, Dictionary<string, object> jobDataMap) where T : IJob
        {
            var jobName = typeof(T).Name;
            IJobDetail jobCheck = JobBuilder.Create<T>().WithIdentity(jobName, jobName + "_Group").Build();
            jobCheck.JobDataMap.PutAll(jobDataMap);
            ISimpleTrigger triggerCheck = new SimpleTriggerImpl(jobName + "_SimpleTrigger",
                jobName + "_TriggerGroup",
                startTime,
                null,
                SimpleTriggerImpl.RepeatIndefinitely,
                simpleTime);
            return await scheduler?.ScheduleJob(jobCheck, triggerCheck);
        }

        /// <summary>
        /// 觸發器1：添加Job並以週期的形式運行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="jobName"></param>
        /// <param name="startTime"></param>
        /// <param name="simpleTime">秒數</param>
        /// <param name="mapKey"></param>
        /// <param name="mapValue"></param>
        /// <returns></returns>
        public static async Task<DateTimeOffset> AddJob<T>(DateTimeOffset startTime, int simpleTime, string mapKey, object mapValue) where T : IJob
        {
            Dictionary<string, object> jobDataMap = new Dictionary<string, object>
            {
                { mapKey, mapValue }
            };
            return await AddJob<T>(startTime, TimeSpan.FromSeconds(simpleTime), jobDataMap);
        }

        /// <summary>
        /// 觸發器1：添加Job並以週期的形式運行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="jobName"></param>
        /// <param name="startTime"></param>
        /// <param name="simpleTime"></param>
        /// <returns></returns>
        public static async Task<DateTimeOffset> AddJob<T>(DateTimeOffset startTime, TimeSpan simpleTime) where T : IJob
        {
            var jobName = typeof(T).Name;
            return await AddJob<T>(startTime, simpleTime, new Dictionary<string, object>());
        }

        /// <summary>
        /// 觸發器1：添加Job並以週期的形式運行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="jobName"></param>
        /// <param name="startTime"></param>
        /// <param name="simpleTime">秒數</param>
        /// <returns></returns>
        public static async Task<DateTimeOffset> AddJob<T>(DateTimeOffset startTime, int simpleTime) where T : IJob
        {
            return await AddJob<T>(startTime, TimeSpan.FromSeconds(simpleTime));
        }

        /// <summary>
        /// 觸發器1：添加Job並以週期的形式運行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="jobName"></param>
        /// <param name="simpleTime">秒數</param>
        /// <returns></returns>
        public static async Task<DateTimeOffset> AddJob<T>(int simpleTime) where T : IJob
        {
            return await AddJob<T>(DateTime.UtcNow.AddSeconds(1), TimeSpan.FromSeconds(simpleTime));
        }
        #endregion

        #region 觸發器4：添加Job
        /// <summary>
        /// 觸發器4：添加Job並以定點的形式運行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="jobName"></param>
        /// <param name="cronTime"></param>
        /// <param name="jobDataMap"></param>
        /// <returns></returns>
        public static async Task<DateTimeOffset> AddJob<T>(string cronTime, string jobData) where T : IJob
        {
            var jobName = typeof(T).Name;
            IJobDetail jobCheck = JobBuilder.Create<T>().WithIdentity(jobName, jobName + "_Group").UsingJobData("jobData", jobData).Build();
            ICronTrigger cronTrigger = new CronTriggerImpl(jobName + "_CronTrigger", jobName + "_TriggerGroup", cronTime);
            return await scheduler?.ScheduleJob(jobCheck, cronTrigger);
        }

        /// <summary>
        /// 觸發器4：添加Job並以定點的形式運行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="jobName"></param>
        /// <param name="cronTime"></param>
        /// <returns></returns>
        public static async Task<DateTimeOffset> AddJob<T>(string cronTime) where T : IJob
        {
            return await AddJob<T>(cronTime, null);
        }
        #endregion

        /// <summary>
        /// 修改觸發器時間重載
        /// </summary>
        /// <param name="jobName">Job名稱</param>
        /// <param name="timeSpan">TimeSpan</param>
        /// </summary>
        public static async Task<bool> UpdateTime<T>(TimeSpan simpleTimeSpan) where T : IJob
        {
            var jobName = typeof(T).Name;
            TriggerKey triggerKey = new TriggerKey(jobName + "_SimpleTrigger", jobName + "_TriggerGroup");
            SimpleTriggerImpl simpleTriggerImpl = (await scheduler?.GetTrigger(triggerKey)) as SimpleTriggerImpl;
            simpleTriggerImpl.RepeatInterval = simpleTimeSpan;
            await scheduler?.RescheduleJob(triggerKey, simpleTriggerImpl);
            return true;
        }

        /// <summary>
        /// 修改觸發器時間重載
        /// </summary>
        /// <param name="jobName">Job名稱</param>
        /// <param name="simpleTime">分鐘數</param>
        /// <summary>
        public static async Task<bool> UpdateTime<T>(int simpleTime) where T : IJob
        {
            var jobName = typeof(T).Name;
            await UpdateTime<T>(TimeSpan.FromMinutes(simpleTime));
            return true;
        }

        /// <summary>
        /// 修改觸發器時間重載
        /// </summary>
        /// <param name="jobName">Job名稱</param>
        /// <param name="cronTime">Cron運算式</param>
        public static async Task<bool> UpdateTime<T>(string cronTime) where T : IJob
        {
            var jobName = typeof(T).Name;
            TriggerKey triggerKey = new TriggerKey(jobName + "_CronTrigger", jobName + "_TriggerGroup");
            CronTriggerImpl cronTriggerImpl = scheduler?.GetTrigger(triggerKey).Result as CronTriggerImpl;
            cronTriggerImpl.CronExpression = new CronExpression(cronTime);
            await scheduler?.RescheduleJob(triggerKey, cronTriggerImpl);
            return true;
        }

        /// <summary>
        /// 暫停所有Job
        /// </summary>
        public static void PauseAll()
        {
            scheduler?.PauseAll();
        }

        /// <summary>
        /// 恢復所有Job
        /// </summary>
        public static void ResumeAll()
        {
            scheduler?.ResumeAll();
        }

        /// <summary>
        /// 刪除指定Job
        /// </summary>
        /// <param name="jobName"></param> 
        public static async Task<bool> Delete<T>() where T : IJob
        {
            var jobName = typeof(T).Name;
            JobKey jobKey = new JobKey(jobName, jobName + "_Group");
            await scheduler?.DeleteJob(jobKey);
            return true;
        }

        /// <summary>
        /// 卸載計時器
        /// </summary>
        /// <param name="isWaitForToComplete">是否等待Job執行完成</param>
        public static void Shutdown(bool isWaitForToComplete)
        {
            scheduler?.Shutdown(isWaitForToComplete);
            scheduler = null;
        }
        public static async Task<bool> Stop<T>() where T : IJob
        {
            var jobName = typeof(T).Name;
            JobKey jobKey = new JobKey(jobName, jobName + "_Group");
            await scheduler?.Interrupt(jobKey);
            return true;
        }

    }
}
