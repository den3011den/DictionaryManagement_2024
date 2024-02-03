using AutoMapper;
using DictionaryManagement_Business.Repository.IRepository;
using DictionaryManagement_DataAccess.Data.IntDB;
using DictionaryManagement_Models.IntDBModels;
using DND.EFCoreWithNoLock.Extensions;
using Microsoft.EntityFrameworkCore;

namespace DictionaryManagement_Business.Repository
{
    public class LogEventRepository : ILogEventRepository
    {
        private readonly IntDBApplicationDbContext _db;
        private readonly IMapper _mapper;
        private readonly ILogEventTypeRepository _logEventTypeRepository;

        public LogEventRepository(IntDBApplicationDbContext db, IMapper mapper, ILogEventTypeRepository logEventTypeRepository)
        {
            _db = db;
            _mapper = mapper;
            _logEventTypeRepository = logEventTypeRepository;
        }

        public async Task<IEnumerable<LogEventDTO>> GetAllByTimeInterval(DateTime? startTime, DateTime? endTime)
        {

            if (startTime == null)
                startTime = DateTime.MinValue;
            if (startTime == null)
                endTime = DateTime.MaxValue;

            var hhh2 = _db.LogEvent
                        .Include("LogEventTypeFK")
                        .Include("UserFK")
                        .Where(u => u.EventTime >= startTime && u.EventTime <= endTime).AsNoTracking().ToListWithNoLock();
            return _mapper.Map<IEnumerable<LogEvent>, IEnumerable<LogEventDTO>>(hhh2);
        }

        public async Task<LogEventDTO?> AddRecord(string logEventTypeName, Guid userId, string oldValue, string newValue, bool isError, string description)
        {
            var logEventTypeDTO = _logEventTypeRepository.GetByName(logEventTypeName).GetAwaiter().GetResult();
            int logEventTypeId = 1;
            if (logEventTypeDTO != null)
                logEventTypeId = logEventTypeDTO.Id;

            var objectToAdd = new LogEvent
            {
                LogEventTypeId = logEventTypeId,
                UserId = userId,
                Description = description.Length > (4000 - 1) ? description.Substring(0, 4000 - 1) : description,
                IsError = isError,
                IsCritical = false,
                IsWarning = false,
                OldValue = oldValue.Length > (200 - 1) ? oldValue.Substring(0, 200 - 1) : oldValue,
                NewValue = newValue.Length > (200 - 1) ? newValue.Substring(0, 200 - 1) : newValue,
                EventTime = DateTime.Now
            };

            var addedLogEvent = _db.LogEvent.Add(objectToAdd);
            _db.SaveChanges();
            return _mapper.Map<LogEvent, LogEventDTO>(addedLogEvent.Entity);
        }

        public async Task ToLog<T>(T? oldObject, T newObject, string logEventTypeName, string prefixString, IAuthorizationRepository _authorizationRepository)
        {

            Guid userId = (await _authorizationRepository.GetCurrentUserDTO()).Id;

            string namestr = prefixString + newObject.ToString().Trim();
            if (oldObject != null)
                if (!((oldObject.ToString()).Trim().ToUpper().Equals(newObject.ToString().Trim().ToUpper())))
                    namestr = namestr + " (" + oldObject.ToString().Trim() + ")";

            string description = logEventTypeName + ": ";

            var propertiesList = typeof(T).GetProperties();

            foreach (var property in propertiesList)
            {
                bool needSkip = true;
                foreach (var customAttribute in property.CustomAttributes)
                {
                    if (customAttribute.AttributeType == typeof(ForLogAttribute))
                    {
                        needSkip = false;
                        break;
                    }
                }

                if (needSkip)
                {
                    continue;
                }

                var propertyAttribute = property.GetCustomAttributes(typeof(ForLogAttribute), false);
                string? namePropertyAttribute = (propertyAttribute[0] as ForLogAttribute) == null ? "" : (propertyAttribute[0] as ForLogAttribute).NameProperty;

                object? oldValue = null;

                if (oldObject != null)
                {
                    oldValue = property.GetValue(oldObject);
                }

                var newValue = property.GetValue(newObject);

                string oldValueString = "";
                string newValueString = "";

                if (oldValue == null)
                {
                    oldValueString = "<Пусто>";
                }
                if (newValue == null)
                {
                    newValueString = "<Пусто>";
                }
                if (oldValue != null)
                {
                    if (oldValue.GetType() == typeof(string))
                        oldValueString = (string)oldValue;
                    else
                        if (oldValue.GetType() == typeof(DateTime))
                        oldValueString = ((DateTime)oldValue).ToString("dd.MM.yyyy HH:mm:ss.fff");
                    else
                        oldValueString = oldValue.ToString();
                }
                if (newValue != null)
                {
                    if (newValue.GetType() == typeof(string))
                        newValueString = (string)newValue;
                    else
                        if (newValue.GetType() == typeof(DateTime))
                        newValueString = ((DateTime)newValue).ToString("dd.MM.yyyy HH:mm:ss.fff");
                    else
                        newValueString = newValue.ToString();
                }

                newValueString = newValueString.ToLower() == "true" ? "Да" : (newValueString.ToLower() == "false" ? "Нет" : newValueString);
                oldValueString = oldValueString.ToLower() == "true" ? "Да" : (oldValueString.ToLower() == "false" ? "Нет" : oldValueString);

                if (!newValueString.Equals(oldValueString))
                {
                    description = namestr + " " + (namePropertyAttribute == null ? "" : namePropertyAttribute);
                    await AddRecord(logEventTypeName: logEventTypeName, userId: userId, oldValue: oldValueString, newValue: newValueString, isError: false, description: description);
                }
            }

        }

    }
}

