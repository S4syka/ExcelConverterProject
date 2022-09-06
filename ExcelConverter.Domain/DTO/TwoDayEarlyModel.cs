using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTO
{
    public class TwoDayEarlyModel
    {
        public string CompanyName { get; set; }

        public IEnumerable<TwoDayEarlyHourModel> Hour { get; set; }

        public IEnumerable<ContractorModel> Contractors { get; set; }
    }
}
