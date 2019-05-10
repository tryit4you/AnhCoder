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
        public async Task<ActionResult<IEnumerable<ProductService>>> GetAllAsyc()
        {
            IEnumerable<ProductService> products = await _productRepository.GetAllsAsyc();
            return Ok(products);
        }
        [HttpGet("{id}")]
        public async Task<ActionResult<ProductService>> GetProductServiceAsyc(string id)
        {
            if (id == null)
            {
                return NotFound();
            }
            ProductService product = await _productRepository.GetAsyc(id);
            if (product == null)
            {
                return NotFound();
            }
            return product;
        }
        // PUT: api/ProductServices/5
        [HttpPut("{id}")]
        public async Task<IActionResult> PutProductService(string id, ProductService productService)
        {
            if (id != productService.Id)
            {
                return BadRequest();
            }

            bool updateState= await _productRepository.UpdateProductAsyc(productService);

            if (updateState)
                return Ok("{success}");
            else
                return Ok("{false}");
        }
        [HttpPost]
        public ActionResult PostProductAsyc(ProductService product)
        {
             _productRepository.CreateAsyc(product);
            return Ok("{success}");
        }       // DELETE: api/ProductServices/5
        [HttpDelete("{id}")]
        public ActionResult DeleteProductService(string id)
        {
             _productRepository.RemoveAsync(id);
         
            return Ok("{success}");
        }
    }
}