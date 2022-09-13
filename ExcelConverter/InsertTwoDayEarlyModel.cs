using ExcelConverter.Domain.DTOs;
using ExcelConverter.Domain.Model;
using ExcelConverter.Repository;
using System.Collections.ObjectModel;

namespace ExcelConverter.App
{
    public class InsertTwoDayEarlyModel
    {
        private UnitOfWork _unitOfWork;

        public TwoDayEarlyExcel TwoDayEarlyExcel { get; private set; }
        public ICollection<TwoDayEarlyHour> TwoDayEarlyHours { get; private set; }
        public ICollection<TwoDayEarlyContractor> TwoDayEarlyContractors { get; private set; }
        public ICollection<ContractorHour> ContractorHours { get; private set; }

        public InsertTwoDayEarlyModel(TwoDayEarlyModel model)
        {
            TwoDayEarlyExcel = new();
            TwoDayEarlyHours = new Collection<TwoDayEarlyHour>();
            TwoDayEarlyContractors = new Collection<TwoDayEarlyContractor>();
            ContractorHours = new Collection<ContractorHour>();
            _unitOfWork = new();

            TwoDayEarlyExcel.CompanyName = model.CompanyName;
            TwoDayEarlyExcel.ExcelDate = DateTime.Now;

            int id = _unitOfWork.TwoDayEarlyExcelsRepository.Insert(TwoDayEarlyExcel);
            TwoDayEarlyExcel.Id = id;

            foreach (var item in model.Hour)
            {
                TwoDayEarlyHours.Add(new TwoDayEarlyHour
                {
                    TwoDayEarlyExcelId = id,
                    Hour = item.Hour,
                    Usage = item.Usage,
                    UnbalanceTwoDays = item.UnbalanceTwoDays,
                    ContractSum = item.ContractSum,
                    ContractSumBuy = item.ContractSumBuy,
                    ContractSumSell = item.ContractSum
                });
            }

            foreach (var item in TwoDayEarlyHours)
            {
                _unitOfWork.TwoDayEarlyHoursRepository.Insert(item);
            }

            foreach (var item in model.Contractors)
            {
                var tempContractor = new TwoDayEarlyContractor
                {
                    TwoDayEarlyExcelId = id,
                    BGCode = item.BGCode
                };
                int contractorId = _unitOfWork.TwoDayEarlyContractorsRepository.Insert(tempContractor);

                tempContractor.Id= contractorId;

                TwoDayEarlyContractors.Add(tempContractor);

                int k = 0;
                foreach (var item2 in item.Price)
                {
                    k++;
                    var tempContractorHour = new ContractorHour
                    {
                        TwoDayEarlyContractorsId = contractorId,
                        Volume = item2,
                        Hour = k
                    };
                    int idContractorHour = _unitOfWork.ContractorHoursRepository.Insert(tempContractorHour);

                    tempContractor.Id = idContractorHour;

                    ContractorHours.Add(tempContractorHour);
                }
            }
        }
    }
}