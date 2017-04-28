using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNCDIExternalProjectManage
{
    class PrizeComparer : IEqualityComparer<Prize>
    {
        public bool Equals(Prize x, Prize y)
        {
            //Check whether the compared objects reference the same data.
            if (Object.ReferenceEquals(x, y)) return true;

            //Check whether any of the compared objects is null.
            if (Object.ReferenceEquals(x, null) || Object.ReferenceEquals(y, null))
                return false;

            //Check whether the products' properties are equal.
            return x.PrizeClassify == y.PrizeClassify && x.Project == y.Project &&
                   x.AwardName == y.AwardName && x.Department == y.Department &&
                   x.Name == y.Name && x.AccountName == y.AccountName &&
                   x.PayYear == y.PayYear;
        }

        public int GetHashCode(Prize obj)
        {
            //Check whether the object is null
            //if (Object.ReferenceEquals(obj, null)) return 0;

            //Get hash code for the Name field if it is not null.
            //int hashName = obj.Name.GetHashCode();

            //int hashAccountName = obj.AccountName.GetHashCode();

            ////Get hash code for the Code field.
            //int hashPrizeClassify = obj.PrizeClassify.GetHashCode();

            //int hashProject = obj.Project.GetHashCode();

            //int hashAward = obj.AwardName.GetHashCode();

            ////int hashId = obj.ID.GetHashCode();

            //int hashDepartment = obj.Department.GetHashCode();

            //int hashDeclareDepartment = obj.DeclareDepartment.GetHashCode();

            //int hashPayYear = obj.PayYear.GetHashCode();

            //int hashPrize = obj.PrizeValue.GetHashCode();

            ////Calculate the hash code for the product.
            //return hashName ^ hashAccountName ^ hashPrizeClassify ^ hashProject ^ hashAward ^ hashDepartment ^
            //       hashDeclareDepartment ^ hashPayYear ^ hashPrize;
            ////return 0;
            return obj.GetHashCode();
        }
    }

    class PrizeDepartmentComparer : IEqualityComparer<Prize>
    {
        public bool Equals(Prize x, Prize y)
        {
            if (Object.ReferenceEquals(x, y)) return true;

            //Check whether any of the compared objects is null.
            if (Object.ReferenceEquals(x, null) || Object.ReferenceEquals(y, null))
                return false;
            return x.Department == y.Department;
        }

        public int GetHashCode(Prize obj)
        {
            return 0;
        }
    }
}
