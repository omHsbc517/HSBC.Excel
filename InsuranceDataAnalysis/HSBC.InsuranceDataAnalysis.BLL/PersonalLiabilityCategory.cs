using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class PersonalLiabilityCategory
    {
        public string LiabilityCategoryCode { get; set; }

        public string LiabilityCategoryName { get; set; }

        public string CategoryCode { get; set; }

        public string CategoryName { get; set; }

        private static List<PersonalLiabilityCategory> lstCategory = new List<PersonalLiabilityCategory>();

        public static List<PersonalLiabilityCategory> LstCategory
        {
            get
            {
                lstCategory.Clear();
                lstCategory.Add(new PersonalLiabilityCategory
                {
                    LiabilityCategoryCode = "100",
                    LiabilityCategoryName = "身故",
                    CategoryCode = "Death",
                    CategoryName = "身故"
                });

                lstCategory.Add(new PersonalLiabilityCategory
                {
                    LiabilityCategoryCode = "100",
                    LiabilityCategoryName = "意外身故",
                    CategoryCode = "ADB",
                    CategoryName = "意外身故"
                });

                lstCategory.Add(new PersonalLiabilityCategory
                {
                    LiabilityCategoryCode = "300",
                    LiabilityCategoryName = "重大疾病",
                    CategoryCode = "MI",
                    CategoryName = "重大疾病"
                });

                lstCategory.Add(new PersonalLiabilityCategory
                {
                    LiabilityCategoryCode = "500",
                    LiabilityCategoryName = "全残",
                    CategoryCode = "TPD",
                    CategoryName = "全残"
                });

                lstCategory.Add(new PersonalLiabilityCategory
                {
                    LiabilityCategoryCode = "700",
                    LiabilityCategoryName = "医疗",
                    CategoryCode = "Medical",
                    CategoryName = "医疗"
                });

                return lstCategory;
            }
        }
    }
}
