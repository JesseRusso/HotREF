using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace HotREF
{
    internal class WaterHeater
    {
        private string make;
        string model;
        string tankTypeCode;
        string tankTypeEng;
        string tankTypeFr;
        string tankVolumeCode;
        string tankVolumeValue;
        string tankVolumeEng;
        string tankVolumeFr;
        string dhwEF;
        string energySourceCode;
        string energySourceEng;
        string energySourceFr;
        bool primary = true;

        public WaterHeater(string manufacturer, string modelNumber, string EF, string impGal, bool electric, bool primary)
        {
            make = manufacturer;
            model = modelNumber;
            dhwEF = EF;
            tankVolumeValue = impGal;
            this.primary = primary;
            if (electric)
            {
                energySourceCode = "1";
                energySourceEng = "Electricity";
                energySourceFr = "Électricité";
            }
            else
            {
                energySourceCode = "2";
                energySourceEng = "Natural gas";
                energySourceFr = "Gaz naturel";
            }
            CheckTankType();
            SetVolume();
        }
        private void SetVolume()
        {
            switch (tankVolumeValue)
            {
                case "0":
                    tankVolumeCode = "7";
                    tankVolumeValue = "0";
                    tankVolumeEng = "Not applicable";
                    tankVolumeFr = "Sans objet";
                    break;

                case "41.6":
                    tankVolumeCode = "4";
                    tankVolumeValue = "189.3001";
                    tankVolumeEng = "189.3 L, 41.6 Imp, 50 US gal";
                    tankVolumeFr = "189.3 L, 41.6 imp, 50 gal ÉU";
                    break;

                case "54.1":
                    tankVolumeCode = "5";
                    tankVolumeValue = "246.0999";
                    tankVolumeEng = "246.1 L, 54.1 Imp, 65 US gal";
                    tankVolumeFr = "246.1 L, 54.1 imp, 65 gal ÉU";
                    break;

                case "66.6":
                    tankVolumeCode = "6";
                    tankVolumeValue = "302.8";
                    tankVolumeEng = "302.8 L, 66.6 Imp, 80 US gal";
                    tankVolumeFr = "302.8 L, 66.6 imp, 80 gal ÉU";
                    break;

                default:
                    tankVolumeCode = "1";
                    tankVolumeValue = Math.Round(System.Convert.ToDouble(tankVolumeValue) * 4.54609,3).ToString();
                    tankVolumeEng = "User specified";
                    tankVolumeFr = "Spécifié par l'utilisateur";
                    break;
            }
        }
        public void CheckTankType()
        {
            switch (energySourceCode)
            {
                case "2":
                    if (System.Convert.ToDouble(tankVolumeValue) > 0)
                    {
                        tankTypeCode = "9";
                        tankTypeEng = "Direct vent (sealed)";
                        tankTypeFr = "à évacuation directe (scellé)";
                    }
                    else
                    {
                        tankTypeCode = "12";
                        tankTypeEng = "Instantaneous (condensing)";
                        tankTypeFr = "Instantané (à condensation)";
                    }
                    break;
                case "1":
                    if (System.Convert.ToDouble(tankVolumeValue) > 0)
                    {
                        tankTypeCode = "2";
                        tankTypeEng = "Conventional tank";
                        tankTypeFr = "Réservoir classique";
                    }
                    else
                    {
                        tankTypeCode = "4";
                        tankTypeEng = "Instantaneous";
                        tankTypeFr = "Chauffage instantané";
                    }
                    break;
            }
        }

        public void AddTank()
        {
            string tankOrder;
            if (primary)
            {
                tankOrder = "Primary";
            }
            else
            {
                tankOrder = "Secondary";
            }
            XElement hw = (from el in CreateProp.newHouse.Descendants("Components").Descendants("HotWater")
                              select el).FirstOrDefault();
            hw.Add(
                new XElement(tankOrder,
                    new XAttribute("hasDrainWaterHeatRecovery", "false"),
                    new XAttribute("insulatingBlanket", "0"),
                    new XAttribute("combinedFlue", "false"),
                    new XAttribute("flueDiameter", "0"),
                    new XAttribute("energyStar", "false"),
                    new XAttribute("ecoEnergy", "false"),
                    new XAttribute("userDefinedPilot", "false"),
                    new XAttribute("connectedUnitsDwhr", "0"),
                    new XElement("EquipmentInformation",
                        new XElement("Manufacturer", make),
                        new XElement("Model", model)),
                    new XElement("EnergySource",
                        new XAttribute("code", energySourceCode),
                        new XElement("English", energySourceEng),
                        new XElement("French", energySourceFr)),
                    new XElement("TankType",
                        new XAttribute("code", tankTypeCode),
                        new XElement("English", tankTypeEng),
                        new XElement("French", tankTypeFr)),
                    new XElement("TankVolume",
                        new XAttribute("code", tankVolumeCode),
                        new XAttribute("value", tankVolumeValue),
                        new XElement("English", tankVolumeEng),
                        new XElement("French", tankVolumeFr)),
                    new XElement("EnergyFactor",
                        new XAttribute("code", "2"),
                        new XAttribute("value", dhwEF),
                        new XAttribute("inputCapacity", "0"),
                        new XAttribute("isUniform", "false"),
                        new XElement("English", "User specified"),
                        new XElement("French", "Spécifié par l'utilisateur")),
                    new XElement("TankLocation",
                        new XAttribute("code", "2"),
                        new XElement("English", "Basement"),
                        new XElement("French", "Sous-sol"))));
            if(energySourceCode == "2")
            {
                hw.Element(tankOrder).SetAttributeValue("pilotEnergy", "0");
            }
        }
    }
}
