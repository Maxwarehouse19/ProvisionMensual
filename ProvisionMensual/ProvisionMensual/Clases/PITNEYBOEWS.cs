using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProvisionMensual.Clases
{
    class PITNEYBOEWS
    {
        public string transactionType { get; set; }

        public float amount { get; set; }

        public string transactionId { get; set; }

        public DateTime transactionDateTime { get; set; }

        public string parcelTrackingNumber { get; set; }

        public string status { get; set; }

        public DateTime statusDate { get; set; }

        public string refundDenialReason { get; set; }

        public string service { get; set; }

        public float zone { get; set; }

        public float weightInOunces { get; set; }

        public string packageType { get; set; }

        public float packageLengthInInches { get; set; }

        public float packageWidthInInches { get; set; }

        public float packageHeightInInches { get; set; }

        public string specialServices { get; set; }

        public string originationAddress { get; set; }

        public string destinationAddress { get; set; }

        public string destinationCountry { get; set; }

        public string adjustmentReason { get; set; }

        public float adjustmentId { get; set; }

        public string refundRequestor { get; set; }

        public float postageBalance { get; set; }

        public string merchantRatePlan { get; set; }

        public string packageIndicator { get; set; }

        public string internationalCountryGroup { get; set; }

        public float dimensionalWeightOz { get; set; }

        public float valueOfGoods { get; set; }

        public string description { get; set; }

        public float inductionPostalCode { get; set; }

        public string customMessage1 { get; set; }

        public string customMessage2 { get; set; }
    }
}
