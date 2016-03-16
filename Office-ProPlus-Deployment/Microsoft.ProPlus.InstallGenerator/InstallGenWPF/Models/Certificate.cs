using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGenerator.Models
{
    public class Certificate : IEquatable<Certificate>
    {
        public string IssuerName { get; set; }

        public string  FriendlyName{ get; set; }

        public string ThumbPrint { get; set; }

        public int Order { get; set; }

        public bool Equals(Certificate other)
        {
            var issuer = IssuerName;
            var friendlyName = FriendlyName;
            var thumbprint = ThumbPrint;
            //if (localId == null) localId = "";
            //if (localName == null) localName = "";

            //var otherId = other.Id;
            //var otherName = other.Name;
            //var otherProductId = other.ProductId;
            //if (otherId == null) otherId = "";
            //if (otherName == null) otherName = "";

            //if (localId.ToLower() == otherId.ToLower() && localProductId == otherProductId)
            //    return true;

            //return false;

            return true;
        }

        //public override int GetHashCode()
        //{
        //    int hashFirstName = Id == null ? 0 : Id.GetHashCode();
        //    int hashLastName = Name == null ? 0 : Name.GetHashCode();

        //    return hashFirstName ^ hashLastName;
        //}

    }
}
