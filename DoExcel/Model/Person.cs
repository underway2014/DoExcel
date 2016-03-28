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
        public string PolicyCode { get; set; }
        public string PayMuch { get; set; }
        public string Name { get; set; }
        public string BirthDay { get; set; }
        public string PhoneNumber { get; set; }
        public string Address { get; set; }
        /// <summary>
        /// 签单人
        /// </summary>
        public string SaleName { get; set; }

        public string getData (int index)
        {
            string res = "";
            switch(index)
            {
                case 0:
                    res = PolicyNumber;
                    break;
                case 1:
                    res = PayMuch;
                    break;
                case 2:
                    res = PayTime;
                    break;
                case 3:
                    res = "";
                    break;
                case 4:
                    res = PolicyCode;
                    break;
                case 5:
                    res = Name;
                    break;
                case 6:
                    res = BirthDay;
                    break;
                case 7:
                    res = PhoneNumber;
                    break;
                case 8:
                    res = Address;
                    break;
                case 9:
                    res = SaleName;
                    break;
                case 10:
                    res = OranizationName;
                    break;
            }
            return res;
        }

        public void setData(string data, int index)
        {
            switch(index)
            {
                case 0:
                    OranizationName = data;
                    break;
                case 1:
                    PayTime = data;
                    break;
                case 2:
                    PolicyNumber = data;
                    break;
                case 3:
                    PolicyCode = data;
                    break;
                case 4:
                    PayMuch = data;
                    break;
                case 5:
                    Name = data;
                    break;
                case 6:
                    BirthDay = data;
                    break;
                case 7:
                    PhoneNumber = data;
                    break;
                case 8:
                    Address = data;
                    break;
                case 9:
                    SaleName = data;
                    break;
                default:
                    break;
            }
        }
    }
}
