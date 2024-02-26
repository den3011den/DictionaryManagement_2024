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

        public async Task<ReportTemplateFileHistoryDTO> Create(ReportTemplateFileHistoryDTO objectToAddDTO)
        {
            if (!String.IsNullOrEmpty(objectToAddDTO.PreviousFileName.Trim()))
            {
                IEnumerable<ReportTemplateFileHistory> previousList = _db.ReportTemplateFileHistory.Where(u => u.CurrentFileName.Trim().ToUpper().Equals(objectToAddDTO.CurrentFileName));
                foreach (ReportTemplateFileHistory item in previousList)
                {
                    item.CurrentFileName = objectToAddDTO.PreviousFileName;
                    _db.Update(item);
                }
            }
            ReportTemplateFileHistory objectToAdd = new ReportTemplateFileHistory();
            objectToAdd.Id = objectToAddDTO.Id;
            objectToAdd.ReportTemplateId = objectToAddDTO.ReportTemplateId;
            objectToAdd.AddTime = objectToAddDTO.AddTime;
            objectToAdd.AddUserId = objectToAddDTO.AddUserId;
            objectToAdd.PreviousFileName = objectToAddDTO.PreviousFileName;
            objectToAdd.CurrentFileName = objectToAddDTO.CurrentFileName;

            var addedReportTemplateFileHistoryEntity = _db.ReportTemplateFileHistory.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<ReportTemplateFileHistory, ReportTemplateFileHistoryDTO>(addedReportTemplateFileHistoryEntity.Entity);
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

