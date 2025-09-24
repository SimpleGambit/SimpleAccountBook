using System.Collections.Generic;
using System.Threading.Tasks;
using SimpleAccountBook.Models;

namespace SimpleAccountBook.Services;

public interface IExcelImportService
{
    Task<IList<TransactionRecord>> LoadAsync(string filePath);
}