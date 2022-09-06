using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DatabaseHelper;
using ExcelConverter.Domain.DTOs;

namespace ExcelConverter.Repository
{
    public class TwoDayEarlyContractorsRepository:BaseRepository<TwoDayEarlyContractor>
    {
        public TwoDayEarlyContractorsRepository(Database database):base(database)
        {

        }
    }
}
