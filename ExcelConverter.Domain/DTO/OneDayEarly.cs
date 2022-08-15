using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTO
{
    public class OneDayEarly
    {
        public IEnumerable<OneDayEarlyHour> Hour { get; private set; }

        public string CompanyName { get; private set; }

        public OneDayEarly(IEnumerable<OneDayEarlyHour> hour, string companyName)
        {
            Hour = hour;
            CompanyName = companyName; 
        }
    }
}
