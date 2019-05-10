using Leaning.Models;
using System;

namespace Leaning
{
   public class Program:CompareInterface<Product>
    {
      
        static void Main(string[] args)
        {
            
            Console.WriteLine("Hello World!");
        }

        public int Comparer(Product entity1, Product entity2)
        {
            throw new NotImplementedException();
        }
    }
}
