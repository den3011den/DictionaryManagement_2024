using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class ReportEntityResendDatesRepository : IReportEntityResendDatesRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public ReportEntityResendDatesRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<IEnumerable<ReportEntityResendDatesDTO>?> GetListByReportEntityId(Guid? reportEntityId)
        {
            if (reportEntityId == null || reportEntityId == Guid.Empty)
            {
                return null;
            }

            var hhh2 = _db.ReportEntityResendDates
                        .Include("ReportEntityFK")
                        .Where(u => u.ReportEntityId == reportEntityId).AsNoTracking().ToListWithNoLock();
            return _mapper.Map<IEnumerable<ReportEntityResendDates>, IEnumerable<ReportEntityResendDatesDTO>>(hhh2);
        }
    }
}

