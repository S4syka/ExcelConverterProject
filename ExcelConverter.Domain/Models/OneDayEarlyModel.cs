using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.Model
{
    public class OneDayEarlyModel
    {
        public IEnumerable<OneDayEarlyHourModel> Hour { get; set; }

        public string CompanyName { get; set; }
    }
}
