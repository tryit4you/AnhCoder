using Leaning.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Leaning
{
    public class Compare : CompareInterface<Product>
    {
        public int Comparer(Product entity1, Product entity2)
        {
            if (entity1.Name.Equals(entity2.Name))
                return 0;
            else if (entity1.Name.Length > entity2.Name.Length)
                return 1;
            else
                return 0;
        }
    }
}
