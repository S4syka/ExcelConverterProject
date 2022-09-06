using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTOs
{
    public class TwoDayEarlyExcel
    {
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public DateTime ExcelDate { get; set; }
        public DateTime CreateDate { get; set; }
        public bool IsDeleted { get; set; }
    }
}
