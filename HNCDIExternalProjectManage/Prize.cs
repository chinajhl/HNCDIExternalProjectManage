using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace HNCDIExternalProjectManage
{
    public class  Prize
    {
        public string Name { get; set; }
        public string AccountName { get; set; }
        public string PrizeClassify { get; set; }
        public string Project { get; set; }
        public string AwardName { get; set; }
        public string PayYear { get; set; }
        public string Department { get; set; }
        public string DeclareDepartment { get; set; }
        public decimal PrizeValue { get; set; }
    }
}
