using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DoExcel.Model
{
    class Person
    {
        /// <summary>
        /// 机构名称
        /// </summary>
        public string OranizationName { get; set; }
        public string PayTime { get; set; }
        /// <summary>
        /// 保单号
        /// </summary>
        public string PolicyNumber { get; set; }
        public string PayMuch { get; set; }
        public string Name { get; set; }
        public string BirthDay { get; set; }
        public string PhoneNumber { get; set; }
        public string Address { get; set; }
        /// <summary>
        /// 签单人
        /// </summary>
        public string SaleName { get; set; }
    }
}
