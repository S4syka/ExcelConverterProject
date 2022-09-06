using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatabaseHelper;


namespace ExcelConverter.Repository
{
    public class UnitOfWork
    {
        private readonly Database _database;

        private readonly Lazy<ContractorHoursRepository> _contractorHoursRepository;
        private readonly Lazy<OneDayEarlyExcelsRepository> _oneDayEarlyExcelsRepository;
        private readonly Lazy<OneDayEarlyHoursRepository> _oneDayEarlyHoursRepository;
        private readonly Lazy<TwoDayEarlyContractorsRepository> _twoDayEarlyContractorsRepository;
        private readonly Lazy<TwoDayEarlyExcelsRepository> _twoDayEarlyExcelsRepository;
        private readonly Lazy<TwoDayEarlyHoursRepository> _twoDayEarlyHoursRepository;

        public UnitOfWork()
        {
            _database = new Database(true);
            _contractorHoursRepository = new Lazy<ContractorHoursRepository>(() => new (_database));
            _oneDayEarlyExcelsRepository = new Lazy<OneDayEarlyExcelsRepository>(() => new (_database));
            _oneDayEarlyHoursRepository = new Lazy<OneDayEarlyHoursRepository>(() => new (_database));
            _twoDayEarlyContractorsRepository = new Lazy<TwoDayEarlyContractorsRepository>(() => new (_database));
            _twoDayEarlyExcelsRepository = new Lazy<TwoDayEarlyExcelsRepository>(() => new (_database));
            _twoDayEarlyHoursRepository = new Lazy<TwoDayEarlyHoursRepository>(() => new(_database));
        }

        public ContractorHoursRepository ContractorHoursRepository => _contractorHoursRepository.Value;
        public OneDayEarlyHoursRepository OneDayEarlyHoursRepository => _oneDayEarlyHoursRepository.Value;
        public OneDayEarlyExcelsRepository OneDayEarlyExcelsRepository => _oneDayEarlyExcelsRepository.Value;
        public TwoDayEarlyContractorsRepository TwoDayEarlyContractorsRepository => _twoDayEarlyContractorsRepository.Value;
        public TwoDayEarlyExcelsRepository TwoDayEarlyExcelsRepository => _twoDayEarlyExcelsRepository.Value;
        public TwoDayEarlyHoursRepository TwoDayEarlyHoursRepository => _twoDayEarlyHoursRepository.Value;

    }
}
