using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatabaseHelper;
using ExcelConverter.Domain.DTOs;

namespace ExcelConverter.Repository
{
    public class ContractorHoursRepository : BaseRepository<ContractorHour>
    {
        public ContractorHoursRepository(Database database): base(database)
        {

        }
    }
}
