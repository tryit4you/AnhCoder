using AnhCoder.Data.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AnhCoder.Data.Interfaces
{
    public interface IProductServiceRepository:IRepository<ProductService>
    {
        Task<bool> UpdateProductAsyc(ProductService product);
    }
}
