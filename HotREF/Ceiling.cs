using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace HotREF
{
    class Ceiling
    {
        string ceilingName;
        string lengthMetric;
        string areaMetric;
        string ceilingType;
        string typeCode;
        string typeEng;
        string typeFr;
        string heelHeight;
        bool vaultCheck = false;
        string ceilingSlope;
        string slopeCode;
        string slopeValue;
        string slopeEng;
        string slopeFr;
        string slopeName = "";
        string rValue;

        public Ceiling(string name, string type, string area, string length, string slope, string heel)
        {
            ceilingName = name;
            ceilingType = type;
            ceilingSlope = slope;
            areaMetric = Math.Round(System.Convert.ToDouble(area) * 0.092903, 4).ToString();
            lengthMetric = Math.Round(System.Convert.ToDouble(length) * 0.3048, 4).ToString();
            heelHeight = Math.Round(System.Convert.ToDouble(heel) * 0.3048, 3).ToString();
            rValue = CreateProp.ceilingRValue;
            SetType();
            SetSlope();
        }

        public Ceiling(string name, string type, string area, string length, string slope,string heel, bool vault)
        {
            vaultCheck = vault;
            ceilingName = name;
            ceilingType = type;
            ceilingSlope = slope;
            heelHeight = Math.Round(System.Convert.ToDouble(heel) * 0.3048, 3).ToString();
            areaMetric = Math.Round(System.Convert.ToDouble(area) * 0.092903, 4).ToString();
            lengthMetric = Math.Round(System.Convert.ToDouble(length) * 0.3048, 4).ToString();
            SetType();
            SetSlope();
        }

        private void SetType()
        {
            switch (ceilingType)
            {
                case "Gable":
                    typeCode = "2";
                    typeEng = "Attic/gable";
                    typeFr = "Combles/pignon";
                    break;
                case "Hip":
                    typeCode = "3";
                    typeEng = "Attic/hip";
                    typeFr = "Combles/arête";
                    break;
                case "Cathedral":
                    typeCode = "4";
                    typeEng = "Cathedral";
                    typeFr = "Cathédrale";
                    break;
                case "Flat":
                    typeCode = "5";
                    typeEng = "Flat";
                    typeFr = "Plat";
                    slopeName = "Flat";
                    break;
                case "Scissor":
                    typeCode = "6";
                    typeEng = "Scissor";
                    typeFr = "Ciseaux";
                    break;
                default:
                    typeCode = "3";
                    typeEng = "Attic/hip";
                    typeFr = "Combles/arête";
                    break;
            }
        }

        private void SetSlope()
        {
            switch (ceilingSlope)
            {
                case "0":
                    slopeCode = "1";
                    slopeValue = "0";
                    slopeEng = "Flat roof";
                    slopeFr = "Toit plat";
                    break;
                case "2":
                    slopeCode = "2";
                    slopeValue = "0.167";
                    slopeEng = "2 / 12";
                    slopeFr = "2 / 12";
                    break;
                case "3":
                    slopeCode = "3";
                    slopeValue = "0.25";
                    slopeEng = "3 / 12";
                    slopeFr = "3 / 12";
                    break;
                case "4":
                    slopeCode = "4";
                    slopeValue = "0.333";
                    slopeEng = "4 / 12";
                    slopeFr = "4 / 12";
                    break;
                case "5":
                    slopeCode = "5";
                    slopeValue = "0.417";
                    slopeEng = "5 / 12";
                    slopeFr = "5 / 12";
                    break;
                case "6":
                    slopeCode = "6";
                    slopeValue = "0.5";
                    slopeEng = "6 / 12";
                    slopeFr = "6 / 12";
                    break;
                case "7":
                    slopeCode = "7";
                    slopeValue = "0.583";
                    slopeEng = "7 / 12";
                    slopeFr = "7 / 12";
                    break;
                default:
                    slopeCode = "0";
                    slopeValue = "0.6667";
                    slopeEng = "User specified";
                    slopeFr = "Spécifié par l'utilisateur";
                    break;
            }
            if (System.Convert.ToInt32(ceilingSlope) > 7)
            {
                slopeCode = "0";
                slopeValue = Math.Round(System.Convert.ToDouble(ceilingSlope)/12,4).ToString();
                slopeEng = "User specified";
                slopeFr = "Spécifié par l'utilisateur";
            }
        }

        public void AddCeiling()
        {
            if (CreateProp.ceilingCount > 1)
            {
                slopeName = ceilingSlope + "/12";
            }
            if (vaultCheck == true)
            {
                rValue = CreateProp.vaultRValue;
                slopeName = "";
            }
            XElement comp = (XElement)(from el in CreateProp.newHouse.Descendants("Components")
                                       select el).First();
            comp.Add(
                new XElement("Ceiling",
                new XAttribute("id", CreateProp.maxID),
                    new XElement("Label", ceilingName + " " + slopeName),
                    new XElement("Construction",
                        new XElement("Type",
                            new XAttribute("code", typeCode),
                            new XElement("English", typeEng),
                            new XElement("French", typeFr)),
                        new XElement("CeilingType", "User specified",
                            new XAttribute("rValue", rValue),
                            new XAttribute("nominalInsulation", rValue))),
                    new XElement("Measurements",
                        new XAttribute("length", lengthMetric),
                        new XAttribute("area", areaMetric),
                        new XAttribute("heelHeight", heelHeight),
                            new XElement("Slope",
                                new XAttribute("code", slopeCode),
                                new XAttribute("value", slopeValue),
                                    new XElement("English", slopeEng),
                                    new XElement("French", slopeFr)))));
            CreateProp.maxID++;
        }
    }
}
