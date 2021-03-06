﻿using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace SZZSNotificationService
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }


        public static void WriteLogToFile(string msg)
        {
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "Log";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            string logPath = AppDomain.CurrentDomain.BaseDirectory + "Log\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";
            try
            {
                using (StreamWriter sw = File.AppendText(logPath))
                {
                    sw.WriteLine("【" + DateTime.Now.ToString() + "】" + msg);
                    sw.Flush();
                    sw.Close();
                    sw.Dispose();
                }
            }
            catch (IOException e)
            {
                using (StreamWriter sw = File.AppendText(logPath))
                {
                    sw.WriteLine("【" + DateTime.Now.ToString() + "】异常：" + e.Message);
                    sw.WriteLine("**************************************************");
                    sw.WriteLine();
                    sw.Flush();
                    sw.Close();
                    sw.Dispose();
                }
            }
        }

        /// <summary>
        /// 删除30天以前产生的日志
        /// </summary>
        public static void DeleteExpiredLogFile()
        {
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "Log";
            if (!Directory.Exists(filePath))
            {
                return;
            }
            string logPath = AppDomain.CurrentDomain.BaseDirectory + "Log\\" + DateTime.Now.AddDays(-30).ToString("yyyy-MM-dd") + ".txt";
            try
            {
                if (File.Exists(logPath))
                {
                    File.Delete(logPath);
                }
            }
            catch (IOException)
            {
                return;
            }
        }

        protected override void OnStart(string[] args)
        {
            StartService();
            /*int resp = 0;
            bool retValue = false;
            int err = 0;
            Interop.ShowMessageBox("This a message from AlertService.", "AlertService Message", out resp, out err, out retValue);
            WriteLogToFile(string.Format("返回值resp={0},err={1},retValue={2}", resp.ToString(), err.ToString(), retValue ? "1" : "0"));*/
        }

        protected override void OnStop()
        {
            this.Dispose();
        }
        private void StartService()
        {
            OnTimedEvent(null, null);
            System.Timers.Timer aTimer = new System.Timers.Timer();
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Interval = 30 * 1000;
            aTimer.Enabled = true;
            GC.KeepAlive(aTimer);
        }

        private static DateTime LastNotificationDate = DateTime.Now.Date.AddDays(-1);
        private static bool isShow = false;
        public void OnTimedEvent(object sender, ElapsedEventArgs e)
        {
            DateTime dateNow = DateTime.Now;
            if (dateNow.Date > LastNotificationDate)
            {
                if (dateNow.Subtract(dateNow.Date.AddHours(9).AddMinutes(30)).Minutes > 0 && dateNow.Subtract(dateNow.Date.AddHours(11).AddMinutes(30)).Minutes < 0)
                {
                    if (GetSZZSResult())
                    {
                        if (!isShow)
                        {
                            isShow = true;
                            int resp = 0;
                            bool retValue = false;
                            int err = 0;
                            Interop.ShowMessageBox("请立即收集晴天卡(确定：立即处理  取消：再次提醒)", "大盘晴雨表", out resp, out err, out retValue);
                            if (resp != 0)
                            {
                                isShow = false;
                                if (resp == 6)
                                {
                                    LastNotificationDate = dateNow.Date;
                                }
                            }
                        }
                    }
                }
                else if (dateNow.Subtract(dateNow.Date.AddHours(13)).Minutes > 0 && dateNow.Subtract(dateNow.Date.AddHours(15)).Minutes < 0)
                {
                    if (GetSZZSResult())
                    {
                        if (!isShow)
                        {
                            isShow = true;
                            int resp = 0;
                            bool retValue = false;
                            int err = 0;
                            Interop.ShowMessageBox("请立即收集晴天卡(确定：立即处理  取消：再次提醒)", "大盘晴雨表", out resp, out err, out retValue);
                            if (resp != 0)
                            {
                                isShow = false;
                                if (resp == 6)
                                {
                                    LastNotificationDate = dateNow.Date;
                                }
                            }
                        }
                    }
                }
            }
        }

        public static bool GetSZZSResult()
        {
            bool retValue = false;
            bool isCollectCard = false;
            if (GetSZZSResultFromEastMoney(ref isCollectCard))
            {
                return isCollectCard;
            }
            else if (GetSZZSResultFrom10jqka(ref isCollectCard))
            {
                return isCollectCard;
            }
            return retValue;
        }

        public static bool GetSZZSResultFromEastMoney(ref bool isCollectCard)
        {
            bool retValue = false;
            try
            {
                HttpClient httpClient = new HttpClient();
                String url = "https://sp0.baidu.com/8aQDcjqpAAV3otqbppnN2DJv/api.php?resource_id=8190&from_mid=1&query=%E4%B8%8A%E8%AF%81%E6%8C%87%E6%95%B0&hilight=disp_data.*.title&sitesign=69fd07e0f24dd2955979fd7abb352f69&eprop=minute";
                HttpResponseMessage response = httpClient.GetAsync(new Uri(url)).Result;
                String s = response.Content.ReadAsStringAsync().Result;
                if (!string.IsNullOrEmpty(s))
                {
                    JObject spropertyObj = JObject.Parse(s);
                    string propertyStr = spropertyObj["data"][0]["disp_data"][0]["property"][0].ToString();
                    JObject propertyObj = JObject.Parse(propertyStr);
                    long _update_time = long.Parse(propertyObj["_update_time"].ToString());
                    long timenow = GetTimeStamp();
                    if (timenow - _update_time >= 60)
                    {
                        return retValue;
                    }
                    string info = propertyObj["data"]["display"]["cur"]["info"].ToString();
                    string currentIndex = info.Substring(info.IndexOf("(") + 1, info.IndexOf(")") - info.IndexOf("(") - 2);
                    double indexVal = 0;
                    if (double.TryParse(currentIndex, out indexVal))
                    {
                        if (indexVal > 0.2)
                        {
                            isCollectCard = true;
                            WriteLogToFile("东方财富网：" + propertyStr);
                        }
                    }
                    retValue = true;
                }
            }
            catch (Exception ex)
            {
                WriteLogToFile("执行方法GetSZZSResultFromEastMoney出现异常，错误详情：" + ex.ToString());
                retValue = false;
            }
            return retValue;
        }

        public static bool GetSZZSResultFrom10jqka(ref bool isCollectCard)
        {
            bool retValue = false;
            try
            {
                HttpClient httpClient = new HttpClient();
                String url = "https://sp0.baidu.com/8aQDcjqpAAV3otqbppnN2DJv/api.php?resource_id=8190&from_mid=1&query=%E4%B8%8A%E8%AF%81%E6%8C%87%E6%95%B0&hilight=disp_data.*.title&sitesign=6d76d21eb8e329b261b4859a72510a9b&eprop=minute";
                HttpResponseMessage response = httpClient.GetAsync(new Uri(url)).Result;
                String s = response.Content.ReadAsStringAsync().Result;
                if (!string.IsNullOrEmpty(s))
                {
                    JObject spropertyObj = JObject.Parse(s);
                    string propertyStr = spropertyObj["data"][0]["disp_data"][0]["property"][0].ToString();
                    JObject propertyObj = JObject.Parse(propertyStr);
                    long _update_time = long.Parse(propertyObj["_update_time"].ToString());
                    long timenow = GetTimeStamp();
                    if (timenow - _update_time >= 60)
                    {
                        return retValue;
                    }
                    string info = propertyObj["data"]["display"]["cur"]["info"].ToString();
                    string currentIndex = info.Substring(info.IndexOf("(") + 1, info.IndexOf(")") - info.IndexOf("(") - 2);
                    double indexVal = 0;
                    if (double.TryParse(currentIndex, out indexVal))
                    {
                        if (indexVal > 0.2)
                        {
                            isCollectCard = true;
                            WriteLogToFile("同花顺：" + propertyStr);
                        }
                    }
                    retValue = true;
                }
            }
            catch (Exception ex)
            {
                WriteLogToFile("执行方法GetSZZSResultFrom10jqka出现异常，错误详情：" + ex.ToString());
                retValue = false;
            }
            return retValue;
        }

        /// <summary> 
        /// 获取时间戳 
        /// </summary> 
        /// <returns></returns> 
        public static long GetTimeStamp()
        {
            TimeSpan ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return long.Parse(Convert.ToInt64(ts.TotalSeconds).ToString());
        }
    }
}
