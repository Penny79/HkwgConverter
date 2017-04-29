using System.Configuration;

namespace HkwgConverter.Model
{
    public class BusinessConfigurationSection : ConfigurationSection
    {

        public BusinessConfigurationSection()
        {

        }

        [ConfigurationProperty("geschaeftsart", DefaultValue = "Intraday-FLEX", IsRequired = true)]
        public string TransactionType
        {
            get
            {
                return (string)this["geschaeftsart"];
            }
            set
            {
                this["geschaeftsart"] = value;
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

        public BusinessPartnerElement()
        {

        }
        [ConfigurationProperty("geschaeftspartnername", IsRequired = true)]
        public string BusinessPartnerName
        {
            get
            {
                return (string)this["geschaeftspartnername"];
            }
            set
            {
                this["geschaeftspartnername"] = value;
            }
        }
    [ConfigurationProperty("ansprechpartner", IsRequired = true)]
        public string ContactPerson {
            get
            {
                return (string)this["ansprechpartner"];
            }
            set
            {
                this["ansprechpartner"] = value;
            }

        }

        [ConfigurationProperty("bilanzkreis", IsRequired = true)]
        public string SettlementArea {
            get
            {
                return (string)this["bilanzkreis"];
            }
            set
            {
                this["bilanzkreis"] = value;
            }
        }
    }
}