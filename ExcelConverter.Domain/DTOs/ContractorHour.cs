using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTOs
{
    public class ContractorHour
    {
        public int Id { get; set; }
        public int TwoDayEarlyContractorId { get; set; }
        public int Hour { get; set; }
        public double Volume { get; set; }
    }
}
