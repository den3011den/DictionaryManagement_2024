using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class ReportTemplateFileHistoryRepository : IReportTemplateFileHistoryRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;

        public ReportTemplateFileHistoryRepository(IntDBApplicationDbContext db, IMapper mapper)
        {
            _db = db;
            _mapper = mapper;
        }

        public async Task<IEnumerable<ReportTemplateFileHistoryDTO>?> GetListByReportEntityId(Guid? reportEntityId)
        {
            if (reportEntityId == null || reportEntityId == Guid.Empty)
            {
                return null;
            }

            var hhh2 = _db.ReportTemplateFileHistory
                        .Include("ReportTemplateFK")
                        .Where(u => u.ReportTemplateId == reportEntityId).AsNoTracking().ToListWithNoLock();
            return _mapper.Map<IEnumerable<ReportTemplateFileHistory>, IEnumerable<ReportTemplateFileHistoryDTO>>(hhh2);
        }
    }
}

