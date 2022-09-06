using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTO
{
    public class ContractorModel
    {
        public string BGCode { get; set; }

        public IEnumerable<double> Price { get; set; }
    }
}
