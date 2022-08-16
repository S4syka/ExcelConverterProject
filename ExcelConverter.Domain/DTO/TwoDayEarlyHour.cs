using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConverter.Domain.DTO
{
    public class TwoDayEarlyHour
    {
        //დრო
        public int Hour { get; set; }

        //მოხმარების პროგნოზი
        public double Usage { get;set; }

        //ორი დღით ადრე უბანალსობა
        public double UnbalanceTwoDays { get; set; }

        //ორმხრივი ხელშეკრულების ჯამი
        public double ContractSum { get; set; }

        //ორმხრივი ხელშეკრულების ჯამი: ყიდვა
        public double ContractSumBuy { get; set; }

        //ორმხრივი ხელშეკრულების ჯამი: გაყიდვა
        public double ContractSumSell { get; set; }
    }
}
