using AnhCoder.Data.Interfaces;
using AnhCoder.Data.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data.Repositories
{
    public class ProductServiceRepository : Repository<ProductService>, IProductServiceRepository
    {
        private readonly AnhCoderDbContext _context;
        public ProductServiceRepository(AnhCoderDbContext context):base(context)
        {
            _context = context;
        }

  
        public async Task<bool> UpdateProductAsyc(ProductService productIn)
        {
            ProductService product = await _context.ProductServices.SingleOrDefaultAsync(x => x.Id == productIn.Id);
            if (product != null)
            {
                product.MetaName = productIn.MetaName;
                product.Description = productIn.Description;
                product.DisCount = productIn.DisCount;
                product.Name = productIn.Name;
                product.PostByUser = productIn.PostByUser;
                product.Price = productIn.Price;
                product.ServiceId = productIn.ServiceId;
                product.Status = productIn.Status;
                product.Thumbnail = productIn.Thumbnail;
                SaveChangeAsyc();
                return true;
            }

            else
            {
                return false;
            }
        }
    }
}
