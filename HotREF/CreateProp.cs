using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace HotREF
{
    public class CreateProp
    {
        private string filePath;
        XDocument newHouse;
        int maxID;
        string wallRValue;

        public CreateProp(String excelFilePath, XDocument template)
        {
            filePath = excelFilePath;
            newHouse = template;
        }

        //Finds the highest value ID attribute
        public void FindID(XDocument newHouse)
        {
            List<string> ids = new List<string>();
            var hasID =
                from el in newHouse.Descendants("House").Descendants().Attributes("id")
                where el.Value != null
                select el.Value;
            foreach (string id in hasID)
            {
                ids.Add(id);
            }

            List<int> idList = ids.Select(s => int.Parse(s)).ToList();
            maxID = idList.Max() + 1;
            ids.Clear();
        }

        public XDocument ChangeWalls()
        {
            XElement first = (XElement)(from el in newHouse.Descendants("Wall")
                                   where el.Element("Label").Value.ToString() == "1st Flr"
                                   select el).First();

            first.Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G21")) * 0.3048, 3));
            first.Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H21")) * 0.3048, 3));
            first.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E21"));
            first.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F21"));
            first.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "K21")) * 0.3048, 3));
            first.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "J21")) * 0.3048, 3));

            XElement second = (XElement)(from el in newHouse.Descendants("Wall")
                                         where el.Element("Label").Value.ToString() == "2nd Flr"
                                         select el).First();

            if (GetCellValue("Calc", "E4") == null || GetCellValue("Calc", "E4") == "0")
            {
                second.Remove();
                first.Descendants("FloorHeader").Remove();
                //change house to 1 storey in specifications if 2nd flr is not present
                foreach (XElement el in newHouse.Descendants("Specifications").Descendants("Storeys"))
                {
                    el.SetAttributeValue("code", "1");
                    el.Element("English").SetValue("One storey");
                    el.Element("French").SetValue("Un étage");
                }
            }
            else
            {
                second.Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G24")) * 0.3048, 3));
                second.Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H24")) * 0.3048, 3));
                second.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E24"));
                second.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F24"));
            }

            XElement garage = (XElement)(from el in newHouse.Descendants("Wall")
                                         where el.Element("Label").Value.ToString() == "Garage Wall"
                                         select el).First();

            if (GetCellValue("Calc", "K3") == null)
            {
                garage.Remove();

                XElement garFloor = (XElement)(from el in newHouse.Descendants("Floor")
                                               where el.Element("Label").Value.Contains("Garage")
                                               select el).First();
                garFloor.Remove();
            }
            else
            {
                garage.Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G30")) * 0.3048, 3));
                garage.Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H30")) * 0.3048, 3));
                garage.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E30"));
                garage.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F30"));
                garage.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "K38")) * 0.3048, 3));
                garage.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "J30")) * 0.3048, 3));
            }

            XElement tall = (XElement)(from el in newHouse.Descendants("Wall")
                                       where el.Element("Label").Value.Equals("Tall Wall")
                                       select el).First();

            if(GetCellValue("Calc", "K4") == null)
            {
                tall.Remove();
            }
            else
            {
                tall.Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G31")) * 0.3048, 3));
                tall.Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H31")) * 0.3048, 3));
                tall.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E31"));
                tall.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F31"));
            }

            XElement plumbing = (XElement)(from el in newHouse.Descendants("Wall")
                                       where el.Element("Label").Value.Contains("Plumbing")
                                       select el).First();

            if (GetCellValue("Calc", "H34") == "0")
            {
                plumbing.Remove();
            }
            else
            {
                plumbing.Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G34")) * 0.3048, 3));
                plumbing.Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H34")) * 0.3048, 3));
                plumbing.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E34"));
                plumbing.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F34"));
            }

            XElement gable = (XElement)(from el in newHouse.Descendants("Wall")
                                           where el.Element("Label").Value.Contains("Gable")
                                           select el).First();

            if (GetCellValue("Calc", "Y18") == "0")
            {
                gable.Remove();
            }
            else
            {
                gable.Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G32")) * 0.3048, 3));
                gable.Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H32")) * 0.3048, 3));
                gable.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E32"));
                gable.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F32"));
            }

            XElement kneewall = (XElement)(from el in newHouse.Descendants("Wall")
                                        where el.Element("Label").Value.Contains("Kneewall")
                                        select el).First();

            if (GetCellValue("Calc", "H29") == "0")
            {
                kneewall.Remove();
            }
            else
            {
                kneewall.Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G29")) * 0.3048, 3));
                kneewall.Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H29")) * 0.3048, 3));
                kneewall.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E29"));
                kneewall.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F29"));
            }
            return newHouse;
        }

        public XDocument ChangeCeilings()
        {
            string type = GetCellValue("Calc", "C10");
            string typeCode;
            string typeEng;
            string typeFr;
            string length = GetCellValue("Calc", "D10");
            string area = GetCellValue("Calc", "E10");
            string slope = GetCellValue("Calc", "F10");
            string slopeCode = "4";
            string slopeValue = "0.333";
            string slopeEng;
            string slopeFr;
            string heel = GetCellValue("Calc", "H10");

            switch (type)
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
                case "Flat":
                    typeCode = "5";
                    typeEng = "Flat";
                    typeFr = "Plat";
                    break;
                default:
                    typeCode = "3";
                    typeEng = "Attic/hip";
                    typeFr = "Combles/arête";
                    break;
            }

            XElement ceil = (XElement)(from el in newHouse.Descendants("Ceiling")
                                        where el.Element("Label").Value.ToString() == "2nd Flr"
                                        select el).First();

            ceil.Element("Construction").Element("Type").SetAttributeValue("code", typeCode);
            ceil.Element("Measurements").SetAttributeValue("length", Math.Round(System.Convert.ToDouble(length)* 0.3048, 3));
            ceil.Element("Measurements").SetAttributeValue("area", Math.Round(System.Convert.ToDouble(area) * 0.3048, 3));
            ceil.Element("Measurements").SetAttributeValue("heelHeight", Math.Round(System.Convert.ToDouble(heel) * 0.3048, 3));

            return newHouse;
        }

        public XDocument ChangeEquipment()
        {
            string furnaceOutput = GetCellValue("General", "C4");
            string hrvMake = GetCellValue("General", "G1");
            string hrvPower1 = GetCellValue("General", "I4");
            string hrvPower2 = GetCellValue("General", "J4");
            string hrvSRE1 = GetCellValue("General", "I5");
            string hrvSRE2 = GetCellValue("General", "J5");
            string hrvFlowrate = GetCellValue("General", "H4");
            string dhwMake = GetCellValue("General", "M1");
            string dhwModel = GetCellValue("General", "N1");
            string dhwSize = GetCellValue("General", "M4");
            string dhwEF = GetCellValue("General", "P4");

            double btus = System.Convert.ToDouble(furnaceOutput);
            double ls = System.Convert.ToDouble(hrvFlowrate);

            // Changes furnace output capacity and EF values
            foreach (XElement furn in newHouse.Descendants("Furnace"))
            {
                furn.Element("Specifications").SetAttributeValue("efficiency", GetCellValue("General", "A5"));
                furn.Element("Specifications").Element("OutputCapacity").SetAttributeValue("value", Math.Round(btus * 0.00029307107, 5).ToString());
                furn.Element("EquipmentInformation").Element("Manufacturer").SetValue(GetCellValue("Summary", "A78"));
                furn.Element("EquipmentInformation").Element("Model").SetValue(GetCellValue("General", "A1"));
            }
            foreach (XElement fan in newHouse.Descendants("HeatingCooling"))
            {
                fan.Element("Type1").Element("FansAndPump").Element("Power").SetAttributeValue("low", GetCellValue("General", "E4"));
                fan.Element("Type1").Element("FansAndPump").Element("Power").SetAttributeValue("high", GetCellValue("General", "D4"));
            }

            //Writes the HRV information
            foreach (XElement hrv in newHouse.Descendants("WholeHouseVentilatorList"))
            {
                hrv.Element("Hrv").Element("EquipmentInformation").Element("Manufacturer").SetValue(hrvMake);
                hrv.Element("Hrv").Element("EquipmentInformation").Element("Model").SetValue(GetCellValue("General", "H1"));
                hrv.Element("Hrv").SetAttributeValue("supplyFlowrate", Math.Round(ls * 0.47195, 4).ToString());
                hrv.Element("Hrv").SetAttributeValue("exhaustFlowrate", Math.Round(ls * 0.47195, 4).ToString());
                hrv.Element("Hrv").SetAttributeValue("fanPower1", hrvPower1);
                hrv.Element("Hrv").SetAttributeValue("fanPower2", hrvPower2);
                hrv.Element("Hrv").SetAttributeValue("efficiency1", hrvSRE1);
                hrv.Element("Hrv").SetAttributeValue("efficiency2", hrvSRE2);
            }

            //Writes the DHW information
            foreach (XElement tank in newHouse.Descendants("Components").Descendants("HotWater"))
            {
                string tankVolumeCode;
                string tankVolumeValue;
                string tankVolumeEng;
                string tankVolumeFr;

                switch (dhwSize)
                {
                    case "50":
                        tankVolumeCode = "4";
                        tankVolumeValue = "189.3001";
                        tankVolumeEng = "189.3 L, 41.6 Imp, 50 US gal";
                        tankVolumeFr = "189.3 L, 41.6 imp, 50 gal ÉU";
                        break;

                    case "65":
                        tankVolumeCode = "5";
                        tankVolumeValue = "246.0999";
                        tankVolumeEng = "246.1 L, 54.1 Imp, 65 US gal";
                        tankVolumeFr = "246.1 L, 54.1 imp, 65 gal ÉU";
                        break;

                    case "80":
                        tankVolumeCode = "6";
                        tankVolumeValue = "302.8";
                        tankVolumeEng = "302.8 L, 66.6 Imp, 80 US gal";
                        tankVolumeFr = "302.8 L, 66.6 imp, 80 gal ÉU";
                        break;

                    default:
                        tankVolumeCode = "1";
                        tankVolumeValue = GetCellValue("General", "O4");
                        tankVolumeEng = "User specified";
                        tankVolumeFr = "Spécifié par l'utilisateur";
                        break;
                }
                tank.Element("Primary").Element("EquipmentInformation").Element("Model").SetValue(dhwModel);
                tank.Element("Primary").Element("EquipmentInformation").Element("Manufacturer").SetValue(dhwMake);
                tank.Element("Primary").Element("TankVolume").SetAttributeValue("code", tankVolumeCode);
                tank.Element("Primary").Element("TankVolume").SetAttributeValue("value", tankVolumeValue);
                tank.Element("Primary").Element("TankVolume").Element("English").SetValue(tankVolumeEng);
                tank.Element("Primary").Element("TankVolume").Element("French").SetValue(tankVolumeFr);
                tank.Element("Primary").Element("EnergyFactor").SetAttributeValue("value", dhwEF);
            }
            return newHouse;
        }
        //Changes values in specifications screen along with the house volume and highest ceiling height
        public XDocument ChangeSpecs()
        {
            string secondFloor = GetCellValue("Calc", "E4");
            string aboveGrade = GetCellValue("Calc", "E6");
            string belowGrade = GetCellValue("Calc", "F6");

            foreach (XElement grade in newHouse.Descendants("Specifications").Descendants("HeatedFloorArea"))
            {
                double areaAbove = System.Convert.ToDouble(aboveGrade);
                double areaBelow = System.Convert.ToDouble(belowGrade);
                areaAbove = Math.Round(areaAbove * 0.092903, 1);
                areaBelow = Math.Round(areaBelow * 0.092903, 1);
                aboveGrade = areaAbove.ToString();
                belowGrade = areaBelow.ToString();

                grade.SetAttributeValue("aboveGrade", aboveGrade);
                grade.SetAttributeValue("belowGrade", belowGrade);
            }

            foreach(XElement infil in newHouse.Descendants("NaturalAirInfiltration"))
            {
                infil.Element("Specifications").Element("House").SetAttributeValue("volume", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "Q49")) * 0.0283168, 3).ToString());
                infil.Element("Specifications").Element("BuildingSite").SetAttributeValue("highestCeiling", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "M51")) * 0.3048, 4).ToString());
            }


            //This section gets the number of corners from each floor and finds the maximum value. Then sets the plan shape 
            List<string> corners = new List<string>();
            corners.Add(GetCellValue("Calc", "F2"));
            corners.Add(GetCellValue("Calc", "F3"));
            corners.Add(GetCellValue("Calc", "F4"));

            corners.RemoveAll(Value => Value == null);

            List<int> shape = new List<int>(corners.Select(s => int.Parse(s)).ToList());
            int maxCorners = shape.Max();
            corners.Clear();

            foreach (XElement ps in newHouse.Descendants("PlanShape"))
            {
                if (maxCorners <= 4)
                {
                    ps.SetAttributeValue("code", "1");
                    ps.Element("English").SetValue("Rectangular");
                    ps.Element("French").SetValue("Rectangulaire");
                }
                if(maxCorners > 4 && maxCorners < 7)
                {
                    ps.SetAttributeValue("code", "4");
                    ps.Element("English").SetValue("Other, 5-6 corners");
                    ps.Element("French").SetValue("Autre, 5-6 coins");
                }
                if (maxCorners > 6 && maxCorners < 9)
                {
                    ps.SetAttributeValue("code", "5");
                    ps.Element("English").SetValue("Other, 7-8 corners");
                    ps.Element("French").SetValue("Autre, 7-8 coins");
                }
                if (maxCorners > 8 && maxCorners < 11)
                {
                    ps.SetAttributeValue("code", "6");
                    ps.Element("English").SetValue("Other, 9-10 corners");
                    ps.Element("French").SetValue("Autre, 9-10 coins");
                }
                if (maxCorners >= 11)
                {
                    ps.SetAttributeValue("code", "7");
                    ps.Element("English").SetValue("Other, 11 or more corners");
                    ps.Element("French").SetValue("Autre, 11 coins ou plus");
                }

            }
        
            return newHouse;
        }
        //Method to get the value of single cells from worksheet
        public string GetCellValue(string sheetName, string refCell)
        {
            string value = null;

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart wbPart = spreadsheet.WorkbookPart;
                    Sheet theSheet = (Sheet)wbPart.Workbook.Descendants<Sheet>().

                        Where(s => s.Name == sheetName).FirstOrDefault();

                    if (theSheet == null)
                    {
                        throw new ArgumentException("sheetName");
                    }
                    WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                    Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                        Where(c => c.CellReference == refCell).FirstOrDefault();

                    if (theCell.CellValue != null)
                    {
                        value = theCell.CellValue.InnerText;

                        if (theCell.DataType != null)
                        {
                            switch (theCell.DataType.Value)
                            {
                                case CellValues.SharedString:

                                    var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                                    if (stringTable != null)
                                    {
                                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                    }
                                    break;

                                case CellValues.Boolean:
                                    switch (value)
                                    {
                                        case "0":
                                            value = "FALSE";
                                            break;
                                        case "1":
                                            value = "TRUE";
                                            break;
                                    }
                                    break;

                                case CellValues.String:
                                    value = theCell.LastChild.InnerText;
                                    break;
                            }

                        }
                    }
                    else
                    {
                        value = null;
                    }
                }
            }
                return value;
        }
        //builds new that are required but not included in the template 
        public XDocument NewWall()
        {
            XElement comp = (XElement)(from el in newHouse.Descendants("Components")
                                       select el).First();

            comp.Add(
                new XElement("Wall",
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", maxID),
                    new XElement("Label", "Test Wall"),
                    new XElement("Construction",
                        new XAttribute("corners", "7"),
                        new XAttribute("intersections", "4"),
                            new XElement("Type", "User specified",
                            new XAttribute("rValue", wallRValue),
                            new XAttribute("nominalInsulation", "3.3527"))),
                    new XElement("Measurements",
                        new XAttribute("height", "0.33"),
                        new XAttribute("perimeter", "12")),
                    new XElement("FacingDirection",
                        new XAttribute("code", "1"),
                        new XElement("English", "N/A"),
                        new XElement("French", "S/O"))));
            maxID++;
            return newHouse;
        }

        public void NewCeiling()
        {

        }
    }
}
