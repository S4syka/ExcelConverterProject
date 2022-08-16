using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTO
{
    public class TwoDayEarly
    {
        public string CompanyName { get; set; }

        public IEnumerable<TwoDayEarlyHour> Hour { get; set; }

        public IEnumerable<Contractor> Contractors { get; set; }
    }
}
