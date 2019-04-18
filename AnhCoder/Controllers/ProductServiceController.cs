using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AnhCoder.Data.Interfaces;
using AnhCoder.Data.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace AnhCoder.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProductServiceController : ControllerBase
    {
        private readonly IProductServiceRepository _productRepository;
        public ProductServiceController(IProductServiceRepository productRepository)
        {
            _productRepository = productRepository;
        }
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ProductService>>> Get()
        {
            IEnumerable<ProductService> products = await _productRepository.GetAllsAsyc();
            return Ok(products);
        }
    }
}