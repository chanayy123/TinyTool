using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
namespace TinyTool
{
    /// <summary>
    /// 考勤时间统计
    /// 白班正常打卡判定时间:07:00-17:00
    /// 夜班正常打卡判定时间:17:00 -07:00
    ///  如果每天第一次打卡时间在早上7点之前判定前一天上夜班,
    ///  如果在17:00之后判定当天是上夜班
    /// </summary>
    class CheckTime
    {
        //迟到时间: 0表示没迟到
        public int LateTime { get; set; }
        //早退时间: 0表示没早退
        public int ExcusedTime { get; set; }
        //加班时间
        public int OverTime { get; set; }
        //工作时间
        public int WorkTime { get; set; }
        //最早打卡时间
        public int FirstTime { get; set; }
        //最晚打卡时间
        public int LastTime { get; set; }
        //新增:是否异常,只有一条打卡记录就判定异常(第一个工作日夜班是可能只有一条打卡记录的)
        public bool IsNormal { get; set; }
        //新增: 是白班还是夜班
        public bool IsDayWork { get; set; }
        //考勤状态:解析打卡时间数据生成, 由考勤补录复写
        public string State { get; set; }
        //考勤状态值
        public string Value { get; set; }
        public CheckTime()
        {
            IsNormal = true;
            IsDayWork = true;
        }
    }
    /// <summary>
    /// 员工当月考勤数据
    /// </summary>
    class StaffCheck
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public DateTime Date { get; set; }
        /// <summary>
        /// key:当月第几天 value:当天所有的打卡纪录
        /// </summary>
        public Dictionary<int, List<DateTime>> RecordTime { get; } = new Dictionary<int, List<DateTime>>();
        /// <summary>
        /// 下个月第一天的考勤记录:为了计算上个月底的夜班时间
        /// </summary>
        public KeyValuePair<int, List<DateTime>> NMFirstDayRecordTime { get; set; }
        /// <summary>
        /// key:当月第几天 value: 当天的考勤具体时间数据
        /// </summary>
        public Dictionary<int, CheckTime> CheckDic { get; set; } = new Dictionary<int, CheckTime>();
    }

    class Program
    {
        public const string SrcXlsPath = "../SrcExcels/";
        public const string DstXlsPath = "../DstExcels/";
        public const string MainSheetName = "总装厂全厂人员";
        public const string RecordSheetName = "考勤补录";
        public const string MainExcelName = "信阳总装11月考勤汇总1-13.xls"; //信阳总装考勤汇总9.1-20
        public const string RollCardExcelName = "9-13.xls";
        public const int MainIDCol = 4;
        public const int MainNameCol = 5;
        public const int MainNameCol2 = 6;
        public const int MainTitleRow = 2;
        public const int MainContentRow = 3;
        //参考上班时间秒数
        public const int StartWorkTime = 8 * 3600;
        //参考下班时间秒数
        public const int EndWorkTime = 17 * 3600 + 1800;
        //参考加班开始时间秒数
        public const int StartOverTime = 18 * 3600;
        //参考上午下班时间
        public const int MorningEndTime = 12 * 3600;
        //参考下午上班时间
        public const int AfternoonStartTime = 13 * 3600 + 1800;
        //参考夜班上班时间秒数
        public const int StartNightWorkTime = 20 * 3600;
        //参考夜班下班时间秒数
        public const int EndNightWorkTime = 29 * 3600;
        //参考夜班加班时间秒数
        public const int StartNightOverTime = 29 * 3600;
        //白班起始时间判定
        public const int DayStartTime = 6 * 3600 + 1800;
        //白班中止时间判定
        public const int DayEndTime = 17 * 3600;

        public const int LastNightStartTime = 19 * 3600;
        public const int LastNightEndTime = 20 * 3600;
        public const int TodayNightEndTime = 8 * 3600;
        public const int TodayNightEndMaxTime = 9 * 3600;

        /// <summary>
        /// key: 员工工号,唯一 value: 员工的考勤纪录
        /// </summary>
        public static Dictionary<string, StaffCheck> AllRecord { get; set; } = new Dictionary<string, StaffCheck>();

        static void Main(string[] args)
        {
            Console.WriteLine("Hello Molly!");
            string totalExcelName = args.Length > 0 ? args[0] : MainExcelName;
            string rollCardExcelName = args.Length > 1 ? args[1] : RollCardExcelName;
            try
            {
                UpdateExcel(totalExcelName, rollCardExcelName);
            }
            catch (Exception e)
            {
                Console.WriteLine("更新表格异常: " + e);
            }
            Console.WriteLine("Bye Molly!");
            //PrintHeart();
            Console.ReadKey();
        }

        static void PrintHeart()
        {
            float y, x, z, f;
            for (y = 1.5f; y > -1.5f; y -= 0.1f)
            {
                for (x = -1.5f; x < 1.5f; x += 0.05f)
                {
                    z = x * x + y * y - 1;
                    f = z * z * z - x * x * y * y * y;
                    Console.Write(f <= 0.0f ? "LOVEMOLLY"[(int)(f * -8.0f)] : ' ');
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// 解析打卡原始数据: 迟到,早退,加班等数据
        /// </summary>
        /// <param name="fileName"></param>
        static void ParseRollCardExcel(string fileName)
        {
            IWorkbook wk = null;
            using (FileStream file = new FileStream(SrcXlsPath + fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                string extension = System.IO.Path.GetExtension(fileName);
                if (extension.Equals(".xls"))
                {
                    //把xls文件中的数据写入wk中
                    wk = new HSSFWorkbook(file, true);
                }
                else
                {
                    //把xlsx文件中的数据写入wk中
                    wk = new XSSFWorkbook(file);
                }
            }
            var sheet = wk.GetSheetAt(0);
            for (var i = 1; i < sheet.PhysicalNumberOfRows; ++i)
            {
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    var id = row.GetCell(1).ToString().Trim();
                    var name = row.GetCell(2).ToString().Trim();
                    var date = row.GetCell(3).DateCellValue;
                    var time = row.GetCell(4).DateCellValue;
                    if (string.IsNullOrEmpty(id)) continue;
                    if (!AllRecord.TryGetValue(id, out StaffCheck staff))
                    {
                        staff = new StaffCheck();
                        staff.Date = date;
                        staff.ID = id;
                        staff.Name = name;
                        List<DateTime> list = new List<DateTime>();
                        list.Add(time);
                        staff.RecordTime.Add(date.Day, list);
                        AllRecord.Add(id, staff);
                    }
                    else
                    {
                        if (date.Month == staff.Date.Month + 1 && date.Day == 1)//下个月考勤
                        {
                            if (staff.NMFirstDayRecordTime.Value == null)
                            {
                                staff.NMFirstDayRecordTime = new KeyValuePair<int, List<DateTime>>(date.Day, new List<DateTime>());
                                staff.NMFirstDayRecordTime.Value.Add(time);
                            }
                            else
                            {
                                staff.NMFirstDayRecordTime.Value.Add(time);
                            }
                        }
                        else if (date.Month == staff.Date.Month)
                        {
                            if (!staff.RecordTime.TryGetValue(date.Day, out List<DateTime> list))
                            {
                                list = new List<DateTime>();
                                staff.RecordTime.Add(date.Day, list);
                            }
                            list.Add(time);
                        }
                        else
                        {
                            Console.WriteLine($"解析{staff.Name}考勤日期异常: {date}");
                        }
                    }
                }
            }
            //所有数据读取完毕，开始解析过滤
            foreach (var item in AllRecord)
            {
                var staff = item.Value;
                foreach (var time in staff.RecordTime)//每一天的刷卡时间都从小到大排序
                {
                    time.Value.Sort((l, r) => CalTotalSeconds(l) - CalTotalSeconds(r));
                    var checkTime = new CheckTime();
                    staff.CheckDic.Add(time.Key, checkTime);
                    //每一天打卡数据只取最早和最晚时间
                    //fix me: 是否休息日目前还没判断,暂定表格里日期添加批注来判断
                    checkTime.IsDayWork = CheckIsDayWork(staff, time.Key);
                    var firstTime = checkTime.IsDayWork ? GetDayWorkFirstTime(staff, time.Key) : GetNightWorkFirstTime(staff, time.Key);
                    var lastTime = checkTime.IsDayWork ? GetDayWorkLastTime(staff, time.Key) : GetNightWorkLastTime(staff, time.Key);
                    if (firstTime == -1 || lastTime == -1)
                    {
                        checkTime.IsNormal = false;
                        Console.WriteLine($"员工 {staff.Name} {time.Key}号 打卡数据异常");
                        continue;
                    }
                    checkTime.FirstTime = firstTime;
                    checkTime.LastTime = lastTime;
                    //统计实际工作时间为了计算休息日加班时长: 目前加班格式不能超过当天24:00
                    //休息日工作时间:需要扣除日常休息时间:中午12:00-13:30,17:30-18:00
                    //超过24点的加班需要人工补录, 要求下班时间要大于上班时间
                    var firstIdx = CheckRollCardTime(ref firstTime, checkTime.IsDayWork);
                    var lastIdx = CheckRollCardTime(ref lastTime, checkTime.IsDayWork);
                    if (checkTime.IsDayWork)
                    {
                        if (firstIdx == lastIdx)
                        {
                            checkTime.WorkTime += (lastTime - firstTime);
                        }
                        else if (firstIdx == 0 && lastIdx == 1)
                        {
                            checkTime.WorkTime += (MorningEndTime - firstTime + lastTime - AfternoonStartTime);
                        }
                        else if (firstIdx == 0 && lastIdx == 2)
                        {
                            checkTime.WorkTime += (MorningEndTime - firstTime + EndWorkTime - AfternoonStartTime + lastTime - StartOverTime);
                        }
                        else if (firstIdx == 1 && lastIdx == 2)
                        {
                            checkTime.WorkTime += (EndWorkTime - firstTime + lastTime - StartOverTime);
                        }
                        //根据修正过的起止刷卡时间
                        //在不同的区间计算真正迟到时间
                        if (firstIdx == 0 && firstTime > StartWorkTime)
                        {
                            checkTime.LateTime = firstTime - StartWorkTime;
                        }
                        else if (firstIdx == 1)
                        {
                            checkTime.LateTime = MorningEndTime - StartWorkTime + firstTime - AfternoonStartTime;
                        }
                        else if (firstIdx == 2)
                        {
                            checkTime.LateTime = MorningEndTime - StartWorkTime + EndWorkTime - AfternoonStartTime + firstTime - StartOverTime;
                        }
                        //在不同的区间计算真正早退时间
                        if (lastIdx == 0)
                        {
                            checkTime.ExcusedTime = MorningEndTime - lastTime + EndWorkTime - AfternoonStartTime;
                        }
                        else if (lastIdx == 1)
                        {
                            checkTime.ExcusedTime = EndWorkTime - lastTime;
                        }
                        //计算普通工作日,包括请假 加班时间,就是从18:00开始算 (休息日加班时间就是工作时间)
                        checkTime.OverTime = checkTime.LastTime - Math.Max(checkTime.FirstTime, StartOverTime);
                    }
                    else
                    {
                        if (firstIdx == 1 && lastIdx == 1)//正常夜班时间段
                        {
                            checkTime.WorkTime += (lastTime - firstTime);
                            if (firstTime > StartNightWorkTime)
                            {
                                checkTime.LateTime = firstTime - StartNightWorkTime;
                            }
                            if (lastTime < EndNightWorkTime)
                            {
                                checkTime.ExcusedTime = EndNightWorkTime - lastTime;
                            }
                            checkTime.OverTime = checkTime.LastTime - Math.Max(checkTime.FirstTime, StartNightOverTime);
                            checkTime.OverTime = Math.Min(checkTime.OverTime, 3*3600);
                        }
                        else //如果在异常时间段,就判定旷工
                        {
                            checkTime.WorkTime += (lastTime - firstTime);
                            checkTime.LateTime = 8 * 3600;
                            checkTime.ExcusedTime = 0;
                            checkTime.OverTime = 0;
                        }

                    }
                }
            }
            Console.WriteLine("打卡数据解析完毕");
        }
        /// <summary>
        /// 新考勤时间
        /// 夜班从晚8点开始到第二天早5点,早5点到早8点算加班时间,最多加班3个小时
        /// 现在夜班时间和白班时间有重叠,之前根据6:30分界线判断已经失效
        /// 仅通过几个特征判断夜班,根据效果再逐步优化: 
        /// 1 当天第一次打卡时间超过8:00
        /// 2 前一天最后打卡时间介于19:00-20:00
        /// </summary>
        /// <param name="check"></param>
        /// <param name="day"></param>
        /// <returns></returns>
        static bool CheckIsDayWork(StaffCheck check, int day)
        {
            check.RecordTime.TryGetValue(day, out List<DateTime> list);
            if (list != null && list.Count > 0)
            {
                check.RecordTime.TryGetValue(day - 1, out List<DateTime> list2);
                if(list2 != null && list2.Count > 0)
                {
                    var s = CalTotalSeconds(list[0]);
                    var s2 = CalTotalSeconds(list2[list2.Count - 1]);
                    if( s >= TodayNightEndTime && s2 >= LastNightStartTime && s2 <= LastNightEndTime)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        static int GetDayWorkFirstTime(StaffCheck check, int day)
        {
            check.RecordTime.TryGetValue(day, out List<DateTime> list);
            if (list != null && list.Count > 0)
            {
                var time = list.Find((t) =>
                {
                    var s = CalTotalSeconds(t);
                    return s >= DayStartTime && s < DayEndTime;
                });
                return CalTotalSeconds(time);
            }
            return -1;
        }

        static int GetDayWorkLastTime(StaffCheck check, int day)
        {
            check.RecordTime.TryGetValue(day, out List<DateTime> list);
            if (list != null && list.Count > 0)
            {
                return CalTotalSeconds(list[list.Count - 1]);
            }
            return -1;
        }

        /// <summary>
        /// 获得员工当天夜班开始时间:从19:00算起第一个有效打卡时间(转换成秒数)
        /// 如果没有获得6:30之前打卡时间
        /// </summary>
        /// <param name="check"></param>
        /// <param name="day"></param>
        /// <returns></returns>
        static int GetNightWorkFirstTime(StaffCheck check, int day)
        {
            check.RecordTime.TryGetValue(day, out List<DateTime> list);
            if (list != null && list.Count > 0)
            {
                var index = list.FindIndex((time) => time.Hour >= 19);
                if (index != -1)
                {
                    return CalTotalSeconds(list[index]);
                }
                //夜班19:00之后打卡数据没有，再次获得6:30之前打卡数据
                index = list.FindIndex((time) =>
                {
                    var s = CalTotalSeconds(time);
                    return s < TodayNightEndTime;
                });
                if (index != -1)
                {
                    return CalTotalSeconds(list[index]);
                }
            }
            return -1;
        }
        /// <summary>
        /// 获取夜班最后一次打卡时间
        /// </summary>
        /// <param name="check"></param>
        /// <param name="day"></param>
        /// <returns></returns>
        static int GetNightWorkLastTime(StaffCheck check, int day)
        {
            //先获取第二天早上6点半之前是否有打卡记录，如果有就是当天夜班的下班时间
            DateTime tmp = new DateTime(check.Date.Year, check.Date.Month, day);
            tmp = tmp.AddDays(1);
            List<DateTime> list1;
            if (tmp.Day == 1)//下个月第一天
            {
                list1 = check.NMFirstDayRecordTime.Value;
            }
            else
            {
                check.RecordTime.TryGetValue(tmp.Day, out list1);
            }
            if (list1 != null && list1.Count > 0)
            {
                var index = list1.FindLastIndex((time) =>
                {
                    var s = CalTotalSeconds(time);
                    return s < TodayNightEndMaxTime;
                });
                if (index != -1)
                {
                    return CalTotalSeconds(24, 0, 0) + CalTotalSeconds(list1[index]);
                }
                else //第二天没有6:30之前打卡记录,再次判定当天晚上有没有打卡记录
                {
                    check.RecordTime.TryGetValue(day, out List<DateTime> list0);
                    if (list0 != null && list0.Count > 0)
                    {
                        index = list0.FindLastIndex((time) => time.Hour >= 19);
                        if (index != -1)
                        {
                            return CalTotalSeconds(list0[index]);
                        }
                        //夜班17:00之后打卡数据没有，再次获得6:30之前打卡数据
                        index = list0.FindIndex((time) =>
                        {
                            var s = CalTotalSeconds(time);
                            return s < TodayNightEndMaxTime;
                        });
                        if (index != -1)
                        {
                            return CalTotalSeconds(list0[index]);
                        }
                    }
                }
            }
            else
            {
                check.RecordTime.TryGetValue(day, out List<DateTime> list0);
                if (list0 != null && list0.Count > 0)
                {
                    var index = list0.FindLastIndex((time) => time.Hour >= 17);
                    if (index != -1)
                    {
                        return CalTotalSeconds(list0[index]);
                    }
                    //夜班17:00之后打卡数据没有，再次获得6:30之前打卡数据
                    index = list0.FindIndex((time) =>
                    {
                        var s = CalTotalSeconds(time);
                        return s < TodayNightEndMaxTime;
                    });
                    if (index != -1)
                    {
                        return CalTotalSeconds(list0[index]);
                    }
                }
            }
            return -1;
        }




        /// <summary>
        /// 修正打卡时间在有效工作区间段,并返回在第几时间段
        /// 白班
        ///  时间段1：8:00-12:00
        ///  时间段2: 13:30-17:30
        ///  时间段3: 18:00-24:00
        ///  夜班
        ///  时间段1 20:00-8:00
        /// </summary>
        /// <param name="seconds"></param>
        /// <returns></returns>
        static int CheckRollCardTime(ref int seconds, bool isDayWork = true)
        {
            var idx = -1;
            if (isDayWork)
            {
                if (seconds < StartWorkTime)//如果打卡时间早于8:00,实际工作时间还是从8:00算
                {
                    seconds = StartWorkTime;
                }
                else if (seconds > MorningEndTime && seconds < AfternoonStartTime)
                {
                    seconds = AfternoonStartTime;
                }
                else if (seconds > EndWorkTime && seconds < StartOverTime)
                {
                    seconds = StartOverTime;
                }
                if (seconds >= StartWorkTime && seconds <= MorningEndTime)
                {
                    idx = 0;
                }
                else if (seconds >= AfternoonStartTime && seconds <= EndWorkTime)
                {
                    idx = 1;
                }
                else if (seconds >= StartOverTime)
                {
                    idx = 2;
                }
            }
            else
            {
                if (seconds >= 19 * 3600 && seconds < StartNightWorkTime)
                {
                    seconds = StartNightWorkTime;
                }
                if (seconds >= StartNightWorkTime)
                {
                    idx = 1;
                }
                else
                {
                    idx = 0;
                }
            }
            return idx;
        }

        static int CalTotalSeconds(DateTime time)
        {
            return time.Hour * 3600 + time.Minute * 60 + time.Second;
        }
        static int CalTotalSeconds(int hour, int minute, int second)
        {
            return hour * 3600 + minute * 60 + second;
        }
        /// <summary>
        /// 计算加班时间,单位:半小时
        /// 不足半小时按 0 算
        /// </summary>
        /// <param name="seconds"></param>
        /// <returns></returns>
        static double CalOverTime(double seconds)
        {
            return Math.Floor(seconds / 1800) * 0.5;
        }
        /// <summary>
        /// 计算迟到早退时间,单位:小时
        ///  3分钟<T<= 1小时: 按一小时算
        ///  1小时<T<= 4小时: 按旷工半天算
        ///  4小时<T 按旷工一天处理
        /// </summary>
        /// <param name="seconds"></param>
        /// <returns></returns>
        static double CalLateExcusedTime(double seconds)
        {
            if (seconds > 3 * 60 && seconds <= 3600)
            {
                seconds = 3600;
            }
            else if (seconds > 3600 && seconds <= 4 * 3600)
            {
                seconds = 4 * 3600;
            }
            else if (seconds > 4 * 3600)
            {
                seconds = 8 * 3600;
            }
            return Math.Floor(seconds / 1800) * 0.5;
        }

        /// <summary>
        /// 先根据打卡数据填充基本考勤纪录, 然后根据考勤补录sheet数据来更新表:考勤补录优先级高会复写字段内容
        /// </summary>
        /// <param name="fileName"></param>
        static void UpdateExcel(string fileName, string rollCardName)
        {
            IWorkbook wk = null;
            using (FileStream file = new FileStream(SrcXlsPath + fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                string extension = System.IO.Path.GetExtension(fileName);
                if (extension.Equals(".xls"))
                {
                    //把xls文件中的数据写入wk中
                    wk = new HSSFWorkbook(file, true);
                }
                else
                {
                    //把xlsx文件中的数据写入wk中
                    wk = new XSSFWorkbook(file);
                }
            }
            var mainSheet = wk.GetSheet(MainSheetName);
            //修正现成加班数据格式:不校验
            //TrimCell(wk);
            //根据打卡数据更新表格
            UpdateExcelByRollCard(rollCardName, mainSheet);
            //获取补录sheet,更新表格
            //UpdateExcelByRecord(wk);
            //当前页内容更新完毕,强制刷新
            //mainSheet.ForceFormulaRecalculation = true;
            WriteToXls(wk, fileName);
        }
        static void UpdateExcelByRollCard(string rollCardFille, ISheet mainSheet)
        {
            ParseRollCardExcel(rollCardFille);
            var curDate = DateTime.Now;
            //根据当前日期来计算之前是否有旷工:如果是空格子,没有打卡数据就认为旷工
            foreach (var item in AllRecord)
            {
                var staff = item.Value;
                for (var i = 1; i <= curDate.Day; ++i)
                {
                    if (staff.CheckDic.ContainsKey(i))
                    {
                        continue;
                    }
                    var rowIdx = GetMainRowIndexByName(mainSheet, staff.Name, staff.ID);
                    var colIdx = GetMainColIndexByDay(mainSheet, i.ToString());
                    if (rowIdx != -1 && colIdx != -1)
                    {
                        var row = mainSheet.GetRow(rowIdx);
                        var cell = row.GetCell(colIdx);
                        if (cell.CellType == CellType.Blank)
                        {
                            var holiday = GetMainHolidayByColIdx(mainSheet, colIdx);
                            cell.SetCellValue(holiday == 0 ? "旷" : "休");
                        }
                    }
                }
            }
            //根据当前每个员工打卡数据更新对应格子考勤状态
            foreach (var item in AllRecord)
            {
                var staff = item.Value;
                foreach (var check in staff.CheckDic)
                {
                    if (!check.Value.IsNormal)
                    {
                        continue;
                    }
                    var rowIdx = GetMainRowIndexByName(mainSheet, staff.Name, staff.ID);
                    var colIdx = GetMainColIndexByDay(mainSheet, check.Key.ToString());
                    if (rowIdx != -1 && colIdx != -1)
                    {
                        var row = mainSheet.GetRow(rowIdx);
                        var row2 = mainSheet.GetRow(rowIdx + 1);
                        var cell = row.GetCell(colIdx);
                        var cell2 = row2.GetCell(colIdx);
                        if (cell == null)
                        {
                            Console.WriteLine($"{staff.Name} {check.Key}号 是空格子!");
                        }
                        //休属于每个月模板内容,如果打卡有休息日加班,应该要复写此类型, 旷是程序生成,一般不会手写,如果有打卡数据了应该也要复写
                        else if (cell.CellType == CellType.Blank || cell.ToString().Contains("休") || cell.ToString().Contains("旷"))
                        {
                            var holiday = GetMainHolidayByColIdx(mainSheet, colIdx);
                            var isHoliday = holiday > 0;
                            if (isHoliday)//休息日的加班时间按实际工作时间计算,休息日没有迟到早退旷工
                            {
                                var overtime = CalOverTime(check.Value.WorkTime);
                                if (overtime > 0 && cell2.ToString() == "")//表里加班数据已填,不校验不覆写
                                {
                                    cell.SetCellValue(check.Value.IsDayWork ? "白" : "夜");
                                    if (holiday == 1)
                                    {
                                        cell2.SetCellValue($"{overtime}(未审)");
                                    }
                                    else
                                    {
                                        cell2.SetCellValue($"{overtime}(未审)");
                                    }
                                    check.Value.State = cell.ToString();
                                }
                            }
                            else if (!isHoliday && cell.ToString().Contains("休"))
                            {
                                Console.WriteLine($"{check.Key}号 不是休息日,却显示 休 !");
                            }
                            else
                            {
                                if (check.Value.ExcusedTime == 0 && check.Value.LateTime == 0) //工作日,正常打卡不迟到不早退
                                {
                                    cell.SetCellValue(check.Value.IsDayWork ? "白" : "夜");
                                    check.Value.State = cell.ToString();
                                }
                                else
                                {
                                    var excusedTime = CalLateExcusedTime(check.Value.ExcusedTime);
                                    var lateTime = CalLateExcusedTime(check.Value.LateTime);
                                    var value = check.Value.IsDayWork ? "白" : "夜";
                                    if (lateTime >= 8 || excusedTime >= 8)
                                    {
                                        value = "旷";
                                    }
                                    else if (lateTime >= 4 && excusedTime >= 4) //迟到导致旷工半天,早退导致旷工半天,累计旷工一天
                                    {
                                        value = "旷";
                                    }
                                    else
                                    {
                                        if (lateTime >= 4)
                                        {
                                            value += "+旷4";
                                        }
                                        else
                                        {
                                            value += lateTime > 0 ? $"+迟{lateTime}" : "";
                                        }
                                        if (excusedTime >= 4)
                                        {
                                            value += "+旷4";
                                        }
                                        else
                                        {
                                            value += excusedTime > 0 ? $"+退{excusedTime}" : "";
                                        }
                                    }
                                    cell.SetCellValue(value);
                                    check.Value.State = cell.ToString();
                                }
                                if (check.Value.OverTime > 0 && cell2.ToString() == "")
                                {
                                    //加班不到半个小时算0
                                    var overtime = CalOverTime(check.Value.OverTime);
                                    if (overtime > 0)
                                    {
                                        cell2.SetCellValue($"{overtime}(未审)");
                                    }
                                }
                            }
                        } //所有请假都需要根据请假时间来重新计算迟到早退数据，目前请假最低半天，上午白天或者下午半天 不会从中间请假
                        //当前考勤数据需要参考格子请假状态才能计算迟到早退,如果格子里没有+,考勤有迟到早退数据, 就会覆盖格子内容,如果有+就不会再覆盖格子: 不管是手动输入+还是程序生成的
                        else if (cell.ToString().StartsWith("事") || cell.ToString().StartsWith("病") || cell.ToString().StartsWith("年") || cell.ToString().StartsWith("调")
                            || cell.ToString().StartsWith("婚") || cell.ToString().StartsWith("伤") || cell.ToString().StartsWith("出") || cell.ToString().StartsWith("产"))
                        {
                            var value = cell.ToString();
                            if (value.Length > 1)//带后缀表明请半天,才有必要计算迟到早退
                            {
                                var time = value.Substring(1);
                                if (!time.Contains("+"))
                                {
                                    var flag = GetWorkTimeArea(check.Value);
                                    int.TryParse(time, out int iTime);
                                    if (iTime != 4)
                                    {
                                        Console.WriteLine($"{staff.Name} {check.Key}号 数据请假格式不对: 比如事/事4");
                                    }
                                    else
                                    {
                                        //半天假,另外半天考勤时间需要重新计算
                                        if (check.Value.IsDayWork)
                                        {
                                            if (flag == 1) //上午请假,下午上班
                                            {
                                                check.Value.LateTime = check.Value.FirstTime > AfternoonStartTime ? check.Value.FirstTime - AfternoonStartTime : 0;
                                                var lastTime = Math.Max(check.Value.LastTime, AfternoonStartTime);//如果忘记打卡,下班时间和上班时间一致，修正一下
                                                check.Value.ExcusedTime = lastTime < EndWorkTime ? EndWorkTime - lastTime : 0;
                                            }
                                            else if (flag == 0) //上午上班，下午请假
                                            {
                                                check.Value.LateTime = check.Value.FirstTime > StartWorkTime ? check.Value.FirstTime - StartWorkTime : 0;
                                                var lastTime = Math.Max(check.Value.LastTime, StartWorkTime);//如果忘记打卡,下班时间和上班时间一致，修正一下
                                                check.Value.ExcusedTime = lastTime < MorningEndTime ? MorningEndTime - lastTime : 0;
                                            }
                                            else if (flag == 2) //半天假, 下班班时间第一次刷卡,旷工半天
                                            {
                                                check.Value.LateTime = EndWorkTime - AfternoonStartTime;
                                                check.Value.ExcusedTime = 0;
                                            }
                                        }
                                        else
                                        {
                                            if (flag == 0)//夜班打卡异常:判定旷工半天
                                            {
                                                check.Value.LateTime = 4 * 3600;
                                                check.Value.ExcusedTime = 0;
                                            }
                                            else //夜班请半天假:就不按上半夜下半夜了，直接计算实际工作时间不够4个小时就算早退
                                            {
                                                if (check.Value.WorkTime < 4 * 3600)
                                                {
                                                    check.Value.LateTime = 0;
                                                    check.Value.ExcusedTime = 4 * 3600 - check.Value.WorkTime;
                                                }
                                                else
                                                {
                                                    check.Value.LateTime = 0;
                                                    check.Value.ExcusedTime = 0;
                                                }
                                            }
                                        }

                                        //重新计算好迟到早退时间,更新请假复合状态
                                        if (check.Value.ExcusedTime != 0 || check.Value.LateTime != 0)
                                        {
                                            var excusedTime = CalLateExcusedTime(check.Value.ExcusedTime);
                                            var lateTime = CalLateExcusedTime(check.Value.LateTime);
                                            var cellValue = cell.ToString();
                                            if (lateTime >= 8 || excusedTime >= 8)//打卡异常情况:请假半天旷工一天
                                            {
                                                cellValue += "+旷(异常:请假半天旷工一天)";
                                            }
                                            else if (lateTime >= 4 && excusedTime >= 4) //打卡异常情况:请假半天旷工一天
                                            {
                                                cellValue += "+旷(异常:请假半天累计旷工一天)";
                                            }
                                            else
                                            {
                                                if (lateTime >= 4)
                                                {
                                                    cellValue += $"+旷4";
                                                }
                                                else
                                                {
                                                    cellValue += lateTime > 0 ? $"+迟{lateTime}" : "";
                                                }
                                                if (excusedTime >= 4)
                                                {
                                                    cellValue += $"+旷4";
                                                }
                                                else
                                                {
                                                    cellValue += excusedTime > 0 ? $"+退{excusedTime}" : "";
                                                }
                                            }
                                            cell.SetCellValue(cellValue);
                                            check.Value.State = cell.ToString();
                                        }
                                        if (check.Value.OverTime > 0 && cell2.ToString() == "")
                                        {
                                            var overtime = CalOverTime(check.Value.OverTime);
                                            if (overtime > 0)//加班以半小时为单位算
                                            {
                                                cell2.SetCellValue($"{overtime}(未审)");
                                            }
                                        }
                                    }
                                }
                            }
                            else //请一天假,只可能计算加班
                            {
                                if (check.Value.OverTime > 0 && cell2.ToString() == "")
                                {
                                    var overtime = CalOverTime(check.Value.OverTime);
                                    if (overtime > 0)//加班以半小时为单位算
                                    {
                                        cell2.SetCellValue($"{overtime}(未审)");
                                    }
                                }
                            }
                        }
                        else //如果当前格子有值,复写打卡数据state,方便以后统计
                        {
                            check.Value.State = cell.ToString();
                            check.Value.Value = cell2.ToString();
                        }
                    }
                }
            }
        }
        static void UpdateExcelByRecord(IWorkbook wk)
        {
            var sheet = wk.GetSheet(RecordSheetName);
            if (sheet == null)
            {
                Console.WriteLine("统计表缺少考勤补录分页!");
            }
            else
            {
                for (var i = 1; i < sheet.PhysicalNumberOfRows; ++i)
                {
                    var row = sheet.GetRow(i);
                    if (row != null)
                    {
                        var id = row.GetCell(0) != null ? row.GetCell(1).ToString() : "";
                        var name = row.GetCell(1).ToString();
                        var day = row.GetCell(2).ToString();
                        var state = row.GetCell(3).ToString();
                        var value = row.GetCell(4) != null ? row.GetCell(4).ToString() : "";
                        if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(day) || string.IsNullOrEmpty(state)) continue;
                        var flag = UpdateMainSheetSingleItem(wk, name, day, state, value, id, true);
                    }
                }
            }
        }
        /// <summary>
        /// 更新考勤明细表:即第二sheet
        /// </summary>
        /// <param name="wk"></param>
        /// <param name="name"></param>
        /// <param name="day"></param>
        /// <param name="state"></param>
        /// <param name="value"></param>
        /// <param name="id"></param>
        /// <param name="reWrite"></param>
        static bool UpdateMainSheetSingleItem(IWorkbook wk, string name, string day, string state, string value, string id, bool reWrite = true)
        {
            try
            {
                ISheet sheet = wk.GetSheet(MainSheetName);
                var rollRecord = GetCheckTimeByName(id, name, day);
                var rowIdx = GetMainRowIndexByName(sheet, name, id);
                var colIdx = GetMainColIndexByDay(sheet, day);
                if (rowIdx != -1 && colIdx != -1)
                {
                    //考勤补录数据同步更新打卡数据,为了统计迟到早退旷工等数据时 还要以state为准
                    if (rollRecord != null)
                    {
                        rollRecord.State = state;
                        rollRecord.Value = value;
                    }
                    var row = sheet.GetRow(rowIdx);
                    var row2 = sheet.GetRow(rowIdx + 1);
                    var cell = row.GetCell(colIdx);
                    var cell2 = row2.GetCell(colIdx);
                    if (reWrite)
                    {
                        if (state == "加")//补录加班数据,还要参考打卡数据，时间要<=打卡时间才合理
                        {
                            var holiday = GetMainHolidayByColIdx(sheet, colIdx);
                            var isHoliday = holiday > 0;
                            if (string.IsNullOrEmpty(value))
                            {
                                Console.WriteLine($"考勤补录: {name} {day}号 加班时间不能为空");
                            }
                            else if (rollRecord != null) //打卡加班数据存在
                            {
                                var overTime = isHoliday ? rollRecord.WorkTime : rollRecord.OverTime;
                                double overtime1 = CalOverTime(overTime);
                                double overtime2 = double.Parse(value);
                                if (overtime2 > overtime1)//申请单时间超过打卡时间不合理
                                {
                                    Console.WriteLine($"考勤补录: {name} {day}号 申请加班时间{overtime2}不合理: 超过打卡加班时间{overtime1}");
                                }
                                else
                                {
                                    if (cell.ToString() == "休" || cell.ToString() == "") //加班会修改下行的加班数值,上行状态如果是休息日或者空,也要修改为白
                                    {
                                        cell.SetCellValue("白");
                                    }
                                    if (holiday == 0)
                                    {
                                        cell2.SetCellValue($"{value}");
                                    }
                                    else if (holiday == 1)
                                    {
                                        cell2.SetCellValue($"{value}");
                                    }
                                    else if (holiday == 2)
                                    {
                                        cell2.SetCellValue($"{value}");
                                    }
                                }
                            }
                            else
                            {
                                if (cell.ToString() == "休" || cell.ToString() == "")//加班会修改下行的加班数值,上行状态如果是休息日或者空,也要修改为白
                                {
                                    cell.SetCellValue("白");
                                }
                                if (holiday == 0)
                                {
                                    cell2.SetCellValue($"{value}");
                                }
                                else if (holiday == 1)
                                {
                                    cell2.SetCellValue($"{value}");
                                }
                                else if (holiday == 2)
                                {
                                    cell2.SetCellValue($"{value}");
                                }
                            }
                        }
                        else
                        {
                            cell.SetCellValue($"{state}{value}");
                        }
                        return true;
                    }
                    else //如果复写为false,只有空白才会写入内容
                    {
                        //if (state == "加" && cell2.CellType == CellType.Blank)//补录加班数据,还要参考打卡数据，时间要<=打卡时间才合理
                        //{
                        //    cell2.SetCellValue(value);
                        //    return true;
                        //}
                        //else if (state != "加" && cell.CellType == CellType.Blank)
                        //{
                        //    cell.SetCellValue($"{state}{value}");
                        //    return true;
                        //}
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("考勤补录: 更新表格错误: " + e);
            }
            return false;
        }

        static CheckTime GetCheckTimeByName(string id, string name, string day)
        {
            //优先判断id存在的情况
            int.TryParse(day, out int iDay);
            if (!string.IsNullOrEmpty(id))
            {
                AllRecord.TryGetValue(id, out StaffCheck staff);
                if (staff == null)
                {
                    Console.WriteLine($"打卡表: 没有此工号 ${id}");
                    return null;
                }
                else
                {
                    staff.CheckDic.TryGetValue(iDay, out CheckTime check);
                    if (check == null)
                    {
                        Console.WriteLine($"打卡表: 工号 {id} {day}号 没有打卡数据");
                        return null;
                    }
                    return check;
                }
            }
            var nameCount = 0;
            var targetId = "";
            //没有填id,就以名字来查找,必须名字保持唯一性
            foreach (var item in AllRecord)
            {
                if (item.Value.Name == name)
                {
                    ++nameCount;
                    targetId = item.Key;
                }
            }
            if (nameCount == 0)
            {
                Console.WriteLine($"打卡表: 没有 ${name}");
                return null;
            }
            if (nameCount > 1)
            {
                Console.WriteLine($"打卡表: 有多个 ${name} ,请输入工号确认员工");
                return null;
            }
            var target = AllRecord[targetId];
            target.CheckDic.TryGetValue(iDay, out CheckTime check2);
            if (check2 == null)
            {
                Console.WriteLine($"打卡表: 姓名 {name} {day}号 没有打卡数据");
                return null;
            }
            return check2;
        }

        static void WriteToXls(IWorkbook wk, string fileName)
        {
            using (FileStream file = new FileStream(DstXlsPath + fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                try
                {
                    wk.Write(file);
                    wk.Close();
                    Console.WriteLine("导出表格完成");
                }
                catch (Exception e)
                {
                    Console.WriteLine("导出表格错误: " + e);
                }

            }
        }
        /// <summary>
        /// 通过格子批注,根据日期获取统计表这一天是否是休息日
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="day"></param>
        /// 返回: 0 表示不是休息日 1 表示休息日 
        /// <returns></returns>
        static int GetMainHolidayByColIdx(ISheet mainSheet, int colIdx)
        {
            var titleRow = mainSheet.GetRow(MainTitleRow);
            var titleCell = titleRow.GetCell(colIdx);
            var comment = titleCell.CellComment?.String.String;
            if (comment != null)
            {
                if (comment.Contains("休"))
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                return 0;
            }
        }

        static int GetMainRowIndexByName(ISheet sheet, string name, string id = "")
        {
            name = name.Trim();
            List<int> rowList = new List<int>();
            for (int i = MainContentRow; i <= sheet.PhysicalNumberOfRows; i++)
            {
                var row = sheet.GetRow(i);  //读取当前行数据
                if (row != null)
                {
                    var nameCell = row.GetCell(MainNameCol);
                    if (nameCell != null)
                    {
                        var nameCell1 = nameCell.ToString().Trim();
                        //var nameCell2 = row.GetCell(MainNameCol2).ToString().Trim();
                        //if (!nameCell1.Equals(nameCell2))
                        //{
                        //    Console.WriteLine($"表格{i + 1}行数据有误:名字不一致!");
                        //    return -1;
                        //}
                        if (nameCell1.Equals(name))
                        {
                            rowList.Add(i);
                        }
                    }
                }
            }
            if (rowList.Count > 1)
            {
                Console.WriteLine($"表格有多条名字 {name} 记录,请确认是否有人同名");
                return -1;
            }
            if (rowList.Count == 0)
            {
                Console.WriteLine($"表格没有此员工: {name} ");
                return -1;
            }
            return rowList[0];
        }

        static int GetMainRowIndexByID(ISheet sheet, string name, string id)
        {
            id = id.Trim();
            List<int> rowList = new List<int>();
            for (int i = MainContentRow; i <= sheet.PhysicalNumberOfRows; i++)
            {
                var row = sheet.GetRow(i);  //读取当前行数据
                if (row != null)
                {
                    var idCell = row.GetCell(MainIDCol).ToString().Trim();
                    if (idCell.Equals(id))
                    {
                        rowList.Add(i);
                    }
                }
            }
            if (rowList.Count > 1)
            {
                Console.WriteLine($"员工 {name} 有多条工号 {id} 记录");
                return -1;
            }
            if (rowList.Count == 0)
            {
                Console.WriteLine($"员工 {name} 缺失ID:{id}");
                return -1;
            }
            return rowList[0];
        }

        static int GetMainColIndexByDay(ISheet sheet, string day)
        {
            day = day.Trim();
            var row = sheet.GetRow(MainTitleRow);
            if (row != null)
            {
                for (var i = 0; i < row.LastCellNum; ++i)
                {
                    var cell = row.GetCell(i);
                    if (cell != null)
                    {
                        var value = cell.ToString().Trim();
                        if (value.Equals(day))
                        {
                            return i;
                        }
                    }
                }
            }
            return -1;
        }
        /// <summary>
        /// 适用请假半天情况: 根据半天打卡数据判断是上午上班还是 下午上班或者 加班时间才上班
        /// 限制: 上午请假半天,打卡时间必须大于12:00
        /// </summary>
        /// <param name="time"></param>
        /// <returns>0:上午上班 1：下午上班 2 加班时间上班</returns>
        static int GetWorkTimeArea(CheckTime time)
        {
            var firstTime = time.FirstTime;
            var firstIdx = CheckRollCardTime(ref firstTime, time.IsDayWork);
            return firstIdx;
        }
        /// <summary>
        /// 根据打卡表修正已经填好的数据
        /// 1 加班数据:之前是纯数字,现在为了区分加班时间类型,需要把纯数字改成 P小时数, X小时数，F小时数
        /// </summary>
        static void TrimCellByRollCard(string mainFilename, string rollCardFilename)
        {
            IWorkbook wk = null;
            using (FileStream file = new FileStream(SrcXlsPath + mainFilename, FileMode.Open, FileAccess.ReadWrite))
            {
                string extension = System.IO.Path.GetExtension(mainFilename);
                if (extension.Equals(".xls"))
                {
                    //把xls文件中的数据写入wk中
                    wk = new HSSFWorkbook(file, true);
                }
                else
                {
                    //把xlsx文件中的数据写入wk中
                    wk = new XSSFWorkbook(file);
                }
            }
            var mainSheet = wk.GetSheet(MainSheetName);
            ParseRollCardExcel(rollCardFilename);
            //解析打卡数据,加班时间从老格式改成新格式
            foreach (var item in AllRecord)
            {
                var staff = item.Value;
                foreach (var check in staff.CheckDic)
                {
                    var rowIdx = GetMainRowIndexByName(mainSheet, staff.Name, staff.ID);
                    var colIdx = GetMainColIndexByDay(mainSheet, check.Key.ToString());
                    if (rowIdx != -1 && colIdx != -1)
                    {
                        var row = mainSheet.GetRow(rowIdx);
                        var row2 = mainSheet.GetRow(rowIdx + 1);
                        var cell = row.GetCell(colIdx);
                        var cell2 = row2.GetCell(colIdx);
                        if (cell2.CellType != CellType.Blank) //加班格子有值,才要复写
                        {
                            var value = cell2.ToString();
                            if (!float.TryParse(value, out float fTime))
                            {
                                Console.WriteLine($"{staff.Name} {check.Key}号 加班数据{value}解析错误!");
                                continue;
                            }
                            var holiday = GetMainHolidayByColIdx(mainSheet, colIdx);
                            var isHoliday = holiday > 0;
                            if (isHoliday)//休息日的加班时间按实际工作时间计算,休息日没有迟到早退旷工
                            {
                                var overtime = CalOverTime(check.Value.WorkTime);
                                if (overtime != fTime)
                                {
                                    var str = holiday == 1 ? "普休" : "法休";
                                    Console.WriteLine($"{staff.Name} {check.Key}号  {str}加班{fTime}小时与考勤{overtime}小时不符");
                                }
                                else
                                {
                                    if (overtime > 0)//加班以半小时为单位算
                                    {
                                        if (holiday == 1)
                                        {
                                            cell2.SetCellValue($"{overtime}");
                                        }
                                        else
                                        {
                                            cell2.SetCellValue($"{overtime}");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (check.Value.OverTime > 0)
                                {
                                    //加班不到半个小时算0
                                    var overtime = CalOverTime(check.Value.OverTime);
                                    if (overtime > 0)
                                    {
                                        if (overtime != fTime)
                                        {
                                            Console.WriteLine($"{staff.Name} {check.Key}号  日常加班数据{fTime}与考勤{overtime}不符");
                                        }
                                        else
                                        {
                                            cell2.SetCellValue($"{overtime}");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            mainSheet.ForceFormulaRecalculation = true;
            WriteToXls(wk, mainFilename);
        }

        /// <summary>
        /// 根据现成表格数据把老格式改成新格式:不做校验
        /// </summary>
        /// <param name="mainFilename"></param>
        /// <param name="rollCardFilename"></param>
        static void TrimCell(IWorkbook wk)
        {
            var mainSheet = wk.GetSheet(MainSheetName);
            //解析打卡数据,加班时间从老格式改成新格式
            var curDate = DateTime.Now;
            var startRow = 3;
            for (var i = startRow; i < mainSheet.PhysicalNumberOfRows; ++i)
            {
                for (var j = 1; j <= curDate.Day; ++j)
                {
                    var row = mainSheet.GetRow(i);
                    if (row != null && row.GetCell(0).CellType == CellType.Blank)//员工考勤第二行
                    {
                        var colIdx = GetMainColIndexByDay(mainSheet, j.ToString());
                        var cell = row.GetCell(colIdx);
                        if (cell != null && cell.CellType != CellType.Blank)
                        {
                            //如果包含+号说明是新格式就不做处理
                            if (cell.ToString().Contains("+")) continue;
                            var holiday = GetMainHolidayByColIdx(mainSheet, colIdx);
                            if (holiday == 0)
                            {
                                cell.SetCellValue($"{cell.ToString()}");
                            }
                            else if (holiday == 1)
                            {
                                cell.SetCellValue($"{cell.ToString()}");
                            }
                            else if (holiday == 2)
                            {
                                cell.SetCellValue($"{cell.ToString()}");
                            }
                        }
                    }
                }
            }
        }
    }
}
