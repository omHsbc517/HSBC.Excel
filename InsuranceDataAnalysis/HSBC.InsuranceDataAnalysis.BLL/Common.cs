﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class Common
    {
        public static string DefaultCommanyName
        {
            get
            {
                return "慕尼黑再保险公司北京分公司";
            }
        }

        public static bool CheckEventType(string yearMonthDay, string effectiveDate)
        {
            effectiveDate = effectiveDate.Replace("-", string.Empty).Substring(0, 6);
            yearMonthDay = yearMonthDay.Substring(0, 6);

            return effectiveDate.Equals(yearMonthDay);
        }

        public static string ConvertToStrToStrDecimal(string value)
        {
            string result = string.Empty;

            result = string.IsNullOrWhiteSpace(value) ? string.Empty :
                        decimal.Parse(value).ToString("0.00");

            return result;
        }

        public static string GetCurrentMonthLastDay(DateTime nowValue)
        {
            string result = string.Empty;
            DateTime d1 = new DateTime(nowValue.Year, nowValue.Month, 1);

            DateTime d2 = d1.AddMonths(1).AddDays(-1);
            result = d2.ToString("yyyy-MM-dd");
            return result;
        }
    }
}