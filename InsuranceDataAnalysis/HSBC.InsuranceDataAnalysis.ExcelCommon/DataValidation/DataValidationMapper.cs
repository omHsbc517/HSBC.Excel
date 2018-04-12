using HSBC.InsuranceDataAnalysis.ExcelCommon.Excel;
using HSBC.InsuranceDataAnalysis.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;

namespace HSBC.InsuranceDataAnalysis.ExcelCommon.DataValidation
{
    public class DataValidationMapper
    {
        private static Dictionary<Type, IEnumerable<ExcelMapping>> getMapper = new Dictionary<Type, IEnumerable<ExcelMapping>>
        {
           { typeof (TEMP_LCProduct),SetMapping<TEMP_LCProduct>()},
           { typeof (TEMP_LCInsured),SetMapping<TEMP_LCInsured>()},
            { typeof (TEMP_LCCont),SetMapping<TEMP_LCCont>()},
             { typeof (TEMP_LCInsureAcc),SetMapping<TEMP_LCInsureAcc>()},
              { typeof (TEMP_LCPolTransaction),SetMapping<TEMP_LCPolTransaction>()}

        };

        internal static IEnumerable<ExcelMapping> GetMapping<T>()
        {
            return getMapper[typeof(T)];
        }

        internal static IEnumerable<ExcelMapping> SetMapping<T>()
        {
            var result = new List<ExcelMapping>();
            var description = "";
            var type=typeof(T);
            var properties = type.GetProperties();
            if (properties!=null)
            {
                foreach (var item in properties)
                {
                    DescriptionAttribute attribute = Attribute.GetCustomAttribute(item, typeof(DescriptionAttribute), false) as DescriptionAttribute;
                    if (attribute!=null)
                    {
                        description = attribute.Description;
                    }
                    result.Add(new ExcelMapping {CoumnName=description,PropertyName=item.Name });
                }
            }
            return result;
        }
    }
}
