using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.OfficeProPlus.InstallGenerator.Models
{
    public class Language : IEquatable<Language>
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string ProductId { get; set; }

        public int Order { get; set; }

        public bool Equals(Language other)
        {
            var localId = Id;
            var localName = Name;
            var localProductId = ProductId;
            if (localId == null) localId = "";
            if (localName == null) localName = "";

            var otherId = other.Id;
            var otherName = other.Name;
            var otherProductId = other.ProductId;
            if (otherId == null) otherId = "";
            if (otherName == null) otherName = "";

            if (localId.ToLower() == otherId.ToLower() && localProductId == otherProductId)
                return true;

            return false;
        }

        public override int GetHashCode()
        {
            int hashFirstName = Id == null ? 0 : Id.GetHashCode();
            int hashLastName = Name == null ? 0 : Name.GetHashCode();

            return hashFirstName ^ hashLastName;
        }

    }
}
