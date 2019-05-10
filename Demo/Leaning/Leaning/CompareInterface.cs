using System;
using System.Collections.Generic;
using System.Text;

namespace Leaning
{
    public interface CompareInterface<TEntity> where TEntity:class
    {
        int Comparer(TEntity entity1, TEntity entity2);
    }
}
