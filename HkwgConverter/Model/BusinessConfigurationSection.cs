using System.Configuration;

namespace HkwgConverter.Model
{
    public class BusinessConfigurationSection : ConfigurationSection
    {
        [ConfigurationProperty("geschäftsart", DefaultValue = "Intraday-FLEX", IsRequired = true)]
        public string TransactionType
        {
            get
            {
                return (string)this["geschäftsart"];
            }
            set
            {
                this["geschäftsart"] = value;
            }
        }

        [ConfigurationProperty("partnerEnviam")]
        public BusinessPartnerElement PartnerEnviaM
        {
            get
            {
                return (BusinessPartnerElement)this["partnerEnviam"];
            }
            set
            {
                this["partnerEnviam"] = value;
            }
        }

        [ConfigurationProperty("partnerCottbus")]
        public BusinessPartnerElement PartnerCottbus
        {
            get
            {
                return (BusinessPartnerElement)this["partnerCottbus"];
            }
            set
            {
                this["partnerCottbus"] = value;
            }
        }
    }

    public class BusinessPartnerElement : ConfigurationElement
    {
        [ConfigurationProperty("geschäftspartnername", DefaultValue = "", IsRequired = true)]
        public string BusinessPartnerName { get; set; }

        [ConfigurationProperty("ansprechpartner", DefaultValue = "", IsRequired = true)]
        public string ContactPerson { get; set; }

        [ConfigurationProperty("bilanzkreis", DefaultValue = "", IsRequired = true)]
        public string SettlementArea { get; set; }
    }
}