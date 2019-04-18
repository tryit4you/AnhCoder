using AnhCoder.Data.Interfaces;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace AnhCoder.Data.Repositories
{
    public class Repository<TEntity> : IRepository<TEntity> where TEntity : class
    {
        protected readonly AnhCoderDbContext _context;
        public Repository(AnhCoderDbContext context)
        {
            _context = context;
        }
        public void CreateAsyc(TEntity entity)
        {
            _context.Set<TEntity>().AddAsync(entity);
        }

        public async Task<IEnumerable<TEntity>> GetAllsAsyc()
        {
            return await _context.Set<TEntity>().ToListAsync();
        }

        public async Task<TEntity> GetAsyc(string id)
        {
           return await _context.Set<TEntity>().FindAsync(id);
        }
       
        public async void RemoveAsync(string id)
        {
            var entity = await GetAsyc(id);
            _context.Set<TEntity>().Remove(entity);
             SaveChangeAsyc();
        }

        public async void SaveChangeAsyc()
        {
            await _context.SaveChangesAsync();
        }

        public async Task Update(string id, TEntity entity)
        {
            _context.Set<TEntity>().Update(entity);
            await _context.SaveChangesAsync();
        }
    }
}
