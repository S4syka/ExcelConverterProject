namespace ExcelConverter.Domain.DTOs
{
    public class TwoDayEarlyHour
    {
        public int Id { get; set; }
        public int TwoDayEarlyExcelId { get; set; }
        public int Hour { get; set; }
        public double Usage { get; set; }   
        public double UnbalanceTwoDays { get; set; }
        public double ContractSum { get; set; }
        public double ContractSumBuy { get; set; }
        public double ContractSumSell { get; set; }
    }
}