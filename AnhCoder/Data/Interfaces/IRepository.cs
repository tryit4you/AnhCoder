using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data.Interfaces
{
    public interface IRepository<TEntity> where TEntity:class
    {
        Task<IEnumerable<TEntity>> GetAllsAsyc();
        Task<TEntity> GetAsyc(string id);
        void RemoveAsync(string id);
        void CreateAsyc(TEntity entity);
        Task Update(string id, TEntity entity);
        void SaveChangeAsyc();
    }
}
