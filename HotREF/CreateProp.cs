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
        public static XDocument newHouse;
        public static int maxID;
        string wallRValue;
        string floorRValue;
        public static string ceilingRValue;
        public static string vaultRValue;
        public static int ceilingCount = 1;

        public CreateProp(String excelFilePath, XDocument template)
        {
            filePath = excelFilePath;
            newHouse = template;
        }

        /// <summary>
        /// Finds the largest component ID in the template and stores it to increment from
        /// </summary>
        /// <param name="newHouse"></param>
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

        public void ChangeWalls()
        {
            XElement first = (XElement)(from el in newHouse.Descendants("Wall")
                                        where el.Element("Label").Value.Contains("1st")
                                        select el).First();

            wallRValue = first.Element("Construction").Element("Type").Attribute("rValue").Value.ToString();

            first.Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G21")) * 0.3048, 3));
            first.Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H21")) * 0.3048, 3));
            first.Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "E21"));
            first.Element("Construction").SetAttributeValue("intersections", GetCellValue("Calc", "F21"));
            first.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "K21")) * 0.3048, 3));
            first.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "J21")) * 0.3048, 3));


            XElement second = (XElement)(from el in newHouse.Descendants("Wall")
                                         where el.Element("Label").Value.Contains("2nd")
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
            GarageWall();
            TallWall();
            PlumbingWall();
        }

        private void GarageWall()
        {
            XElement garage = (XElement)(from el in newHouse.Descendants("Wall")
                                         where el.Element("Label").Value.ToLower().Contains("garage")
                                         select el).First();

            if (GetCellValue("Calc", "K3") == null)
            {
                garage.Remove();
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
        }

        private void TallWall()
        {
            XElement tall = (XElement)(from el in newHouse.Descendants("Wall")
                                       where el.Element("Label").Value.ToLower().Contains("tall")
                                       select el).First();

            if (GetCellValue("Calc", "K4") == null)
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
        }
        private void PlumbingWall()
        {
            XElement plumbing = (XElement)(from el in newHouse.Descendants("Wall")
                                           where el.Element("Label").Value.ToLower().Contains("plumbing")
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
        }
        public void ChangeBasment()
        {
            bool bsmtUnder4Feet = false;
            bool slabOnGrade = false;

            if (GetCellValue("Calc", "N4") == "y" || GetCellValue("Calc", "N4") == "Y")
            {
                bsmtUnder4Feet = true;
            }
            if (GetCellValue("Calc", "N5") == "y" || GetCellValue("Calc", "N5") == "Y")
            {
                slabOnGrade = true;
            }
            XElement over4 = (XElement)(from el in newHouse.Descendants("Components").Descendants("Basement")
                                        where el.Element("Label").Value.Contains(">")
                                        select el).First();

            over4.SetAttributeValue("exposedSurfacePerimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "E38")) * 0.3048, 4));
            over4.Element("Floor").Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "D38")) * 0.3048, 4));
            over4.Element("Floor").Element("Measurements").SetAttributeValue("area", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "F38")) * 0.092903, 4));
            over4.Element("Wall").Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G38")) * 0.3048, 4));
            over4.Element("Wall").Element("Measurements").SetAttributeValue("depth", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H38")) * 0.3048, 4));
            over4.Element("Wall").Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "J38"));
            over4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "K38")) * 0.3048, 4));
            over4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "L38")) * 0.3048, 4));

            if (System.Convert.ToDouble(GetCellValue("Calc", "I38")) > 1)
            {
                over4.Element("Wall").SetAttributeValue("hasPonyWall", "true");
                over4.Element("Wall").Element("Measurements").SetAttributeValue("ponyWallHeight", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "I38")) * 0.3048, 4));
                over4.Element("Wall").Element("Construction").Add(
                    new XElement("PonyWallType",
                        new XAttribute("nominalInsulation", "3.2536"),
                        new XElement("Description", "User specified"),
                        new XElement("Composite",
                            new XElement("Section",
                                new XAttribute("rank", "1"),
                                new XAttribute("percentage", "100"),
                                new XAttribute("rsi", "2.6029"),
                                new XAttribute("nominalRsi", "3.2536")))));
            }
            else
            {
                over4.Element("Wall").SetAttributeValue("hasPonyWall", "false");
                over4.Element("Wall").Element("Construction").Element("PonyWallType").Remove();
                over4.Element("Wall").Element("Measurements").SetAttributeValue("ponyWallHeight", "0");
            }
            Under4Bsmt(bsmtUnder4Feet);
            SlabOnGrade(slabOnGrade);
        }

        private void Under4Bsmt(bool under4Present)
        {
            bool check = under4Present;
            XElement under4 = (XElement)(from el in newHouse.Descendants("Components").Descendants("Basement")
                                         where el.Element("Label").Value.Contains("<")
                                         select el).First();
            if (check == true)
            {
                under4.SetAttributeValue("exposedSurfacePerimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "E39")) * 0.3048, 4));
                under4.Element("Floor").Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "D39")) * 0.3048, 4));
                under4.Element("Floor").Element("Measurements").SetAttributeValue("area", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "F39")) * 0.092903, 4));
                under4.Element("Wall").Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "G39")) * 0.3048, 4));
                under4.Element("Wall").Element("Measurements").SetAttributeValue("depth", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "H39")) * 0.3048, 4));
                under4.Element("Wall").Element("Construction").SetAttributeValue("corners", GetCellValue("Calc", "J39"));
                under4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("height", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "K39")) * 0.3048, 4));
                under4.Element("Components").Element("FloorHeader").Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "L39")) * 0.3048, 4));

                under4.Element("Wall").SetAttributeValue("hasPonyWall", "true");
                under4.Element("Wall").Element("Measurements").SetAttributeValue("ponyWallHeight", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "I39")) * 0.3048, 4));
            }
            else
            {
                under4.Remove();
            }
        }
        private void SlabOnGrade(bool slabPresent)
        {
            bool check = slabPresent;
            XElement slab = (XElement)(from el in newHouse.Descendants("Components").Descendants("Slab")
                                       where el.Element("Label").Value.Contains("Slab")
                                       select el).First();
            if(check == true)
            {
                slab.SetAttributeValue("exposedSurfacePerimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "E40")) * 0.3048, 4));
                slab.Element("Floor").Element("Measurements").SetAttributeValue("area", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "F40")) * 0.092903, 4));
                slab.Element("Floor").Element("Measurements").SetAttributeValue("perimeter", Math.Round(System.Convert.ToDouble(GetCellValue("Calc", "D40")) * 0.3048, 4));
            }
            else
            {
                slab.Remove();
            }
        }
        /// <summary>
        /// Gathers R values from existing ceiling elements in the H2K template
        /// </summary>
        public void CheckCeilings()
        {
            XElement ceil = (XElement)(from el in newHouse.Descendants("Ceiling")
                                       where el.Element("Label").Value.Contains("2nd")
                                       select el).First();

            ceilingRValue = ceil.Element("Construction").Element("CeilingType").Attribute("rValue").Value.ToString();
            ceil.Remove();

            XElement vault = (XElement)(from el in newHouse.Descendants("Components").Descendants("Ceiling")
                                        where el.Element("Construction").Element("Type").Attribute("code").Value == "6"
                                        select el).First();

            vaultRValue = vault.Element("Construction").Element("CeilingType").Attribute("rValue").Value;
            vault.Remove();                                         
        }

        public XDocument ChangeEquipment()
        {
            string furnaceModel = GetCellValue("Summary", "B78");
            string furnaceOutput = GetCellValue("General", "C4");
            string hrvMake = GetCellValue("Summary", "D74");
            string hrvModel = GetCellValue("Summary", "E74");
            string hrvPower1 = GetCellValue("General", "I4");
            string hrvPower2 = GetCellValue("General", "J4");
            string hrvSRE1 = GetCellValue("General", "I5");
            string hrvSRE2 = GetCellValue("General", "J5");
            string hrvFlowrate = GetCellValue("General", "H4");
            string dhwMake = GetCellValue("Summary", "J74");
            string dhwModel = GetCellValue("Summary", "K74");
            string dhwSize = GetCellValue("Summary", "K75");
            string dhwEF = GetCellValue("Summary", "K77");

            double btus = System.Convert.ToDouble(furnaceOutput);

            if (Convert.ToDouble(GetCellValue("General", "B4")) > 0)
            {
                furnaceModel += " " + "& " + GetCellValue("Summary", "B79");
            }
            // Changes furnace output capacity and EF values
            foreach (XElement furn in newHouse.Descendants("Furnace"))
            {
                furn.Element("Specifications").SetAttributeValue("efficiency", GetCellValue("General", "A5"));
                furn.Element("Specifications").Element("OutputCapacity").SetAttributeValue("value", Math.Round(btus * 0.00029307107, 5).ToString());
                furn.Element("EquipmentInformation").Element("Manufacturer").SetValue(GetCellValue("Summary", "A78"));
                furn.Element("EquipmentInformation").Element("Model").SetValue(furnaceModel);
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
                hrv.Element("Hrv").Element("EquipmentInformation").Element("Model").SetValue(hrvModel);
                hrv.Element("Hrv").SetAttributeValue("supplyFlowrate", Math.Round(System.Convert.ToDouble(hrvFlowrate) * 0.47195, 4).ToString());
                hrv.Element("Hrv").SetAttributeValue("exhaustFlowrate", Math.Round(System.Convert.ToDouble(hrvFlowrate) * 0.47195, 4).ToString());
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
                string tankTypeCode = null;
                string tankTypeEng = "Direct vent (sealed)";
                string tankTypeFr = "à évacuation directe (scellé)";

                switch (dhwSize)
                {
                    case "0":
                        tankTypeCode = "12";
                        tankTypeEng = "Instantaneous (condensing)";
                        tankTypeFr = "Instantané (à condensation)";
                        tankVolumeCode = "7";
                        tankVolumeValue = "0";
                        tankVolumeEng = "Not applicable";
                        tankVolumeFr = "Sans objet";
                        break;

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

                if (tankTypeCode != null)
                {
                    tank.Element("Primary").Element("TankType").SetAttributeValue("code", tankTypeCode);
                    tank.Element("Primary").Element("TankType").Element("English").SetValue(tankTypeEng);
                    tank.Element("Primary").Element("TankType").Element("French").SetValue(tankTypeFr);
                }
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

            foreach (XElement infil in newHouse.Descendants("NaturalAirInfiltration"))
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
                if (maxCorners > 4 && maxCorners < 7)
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
        public XDocument ChangeFloors()
        {
            double garFlrArea = System.Convert.ToDouble(GetCellValue("Calc", "P21"));
            double garFlrLength = System.Convert.ToDouble(GetCellValue("Calc", "O21"));

            XElement garFlr = (XElement)(from el in newHouse.Descendants("Floor")
                                         where el.Element("Label").Value.ToLower().Contains("garage")
                                         select el).First();

            XElement floor = (XElement)(from el in newHouse.Descendants("Floor")
                                        where el.Element("Label").Value.Contains("Cant")
                                        select el).First();

            floorRValue = floor.Element("Construction").Element("Type").Attribute("rValue").Value.ToString();
            floor.Remove();

            if ((GetCellValue("Calc", "P21") != null) && (double.Parse(GetCellValue("Calc", "P21")) > 0))
            {
                garFlr.Element("Measurements").SetAttributeValue("area", Math.Round(garFlrArea * 0.092903, 4));
                garFlr.Element("Measurements").SetAttributeValue("length", Math.Round(garFlrLength * 0.3048, 4));
                garFlr.Element("Label").SetValue(GetCellValue("Calc", "L21"));
            }
            else
            {
                garFlr.Remove();
            }
            return newHouse;
        }
        public void ExtraFloors()
        {
            string column = "P";
            int startRow = 22;
            int endRow = 34;
            int currentRow = startRow;
            string area;
            string length;
            string name;

            for (int i = startRow; i <= endRow; i++)
            {
                string currentCell = column + currentRow.ToString();
                if ((GetCellValue("Calc", currentCell) != null) && double.Parse(GetCellValue("Calc", currentCell)) > 0)
                {
                    area = GetCellValue("Calc", currentCell);
                    length = GetCellValue("Calc", "O" + currentRow);
                    name = GetCellValue("Calc", "L" + currentRow);

                    NewFloor(name, length, area);
                }
                currentRow++;
            }
        }
        public void ExtraCeilings()
        {
            string column = "E";
            int startRow = 10;
            int endRow = 17;
            int currentRow = startRow;
            string type;
            string length;
            string area;
            string slope;
            string heel;
            string name;
            List<Ceiling> ceilings = new List<Ceiling>();

            for (int i = startRow; i <= endRow; i++)
            {
                string currentCell = column + currentRow.ToString();
                if ((GetCellValue("Calc", currentCell) != null) && double.Parse(GetCellValue("Calc", currentCell)) > 0)
                {
                    area = GetCellValue("Calc", currentCell);
                    length = GetCellValue("Calc", "D" + currentRow);
                    name = GetCellValue("Calc", "A" + currentRow);
                    type = GetCellValue("Calc", "C" + currentRow);
                    slope = GetCellValue("Calc", "F" + currentRow);
                    heel = GetCellValue("Calc", "H" + currentRow);
                    ceilings.Add(new Ceiling(name, type, area, length, slope, heel));
                    ceilingCount = ceilings.Count();
                }
                currentRow++;
            }
            ceilingCount = ceilings.Count();
            foreach (Ceiling c in ceilings)
            {
                c.AddCeiling();
            }
        }

        public void CheckVaults()
        {
            string column = "M";
            int startRow = 10;
            int endRow = 17;
            int currentRow = startRow;
            string type;
            string length;
            string area;
            string slope;
            string name;
            string heel = GetCellValue("Calc", "H10");
            List<Ceiling> vaults = new List<Ceiling>();

            for (int i = startRow; i <= endRow; i++)
            {
                string currentCell = column + currentRow.ToString();
                if ((GetCellValue("Calc", currentCell) != null) && double.Parse(GetCellValue("Calc", currentCell)) > 0)
                {
                    area = GetCellValue("Calc", currentCell);
                    length = GetCellValue("Calc", "L" + currentRow);
                    name = GetCellValue("Calc", "I" + currentRow);
                    type = GetCellValue("Calc", "AD" + currentRow);
                    slope = GetCellValue("Calc", "N" + currentRow);
                    vaults.Add(new Ceiling(name, type, area, length, slope, heel, true));
                }
                currentRow++;
            }
            foreach (Ceiling v in vaults)
            {
                v.AddCeiling();
            }
        }

        //builds new walls that are required but not included in the template 
        public XDocument NewWall(string name, string corners, string intersections, string height, string perim)
        {
            string heightMetric = Math.Round(System.Convert.ToDouble(height) * 0.3048, 4).ToString();
            string perimMetric = Math.Round(System.Convert.ToDouble(perim) * 0.3048, 4).ToString();
            XElement comp = (XElement)(from el in newHouse.Descendants("Components")
                                       select el).First();

            comp.Add(
                new XElement("Wall",
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", maxID),
                    new XElement("Label", name),
                    new XElement("Construction",
                        new XAttribute("corners", corners),
                        new XAttribute("intersections", intersections),
                            new XElement("Type", "User specified",
                            new XAttribute("rValue", wallRValue),
                            new XAttribute("nominalInsulation", "3.3527"))),
                    new XElement("Measurements",
                        new XAttribute("height", heightMetric),
                        new XAttribute("perimeter", perimMetric)),
                    new XElement("FacingDirection",
                        new XAttribute("code", "1"),
                        new XElement("English", "N/A"),
                        new XElement("French", "S/O"))));
            maxID++;
            return newHouse;
        }

        public XDocument NewFloor(string name, string length, string area)
        {
            string lengthMetric = Math.Round(System.Convert.ToDouble(length) * 0.3048, 4).ToString();
            string areaMetric = Math.Round(System.Convert.ToDouble(area) * 0.092903, 4).ToString();

            XElement comp = (XElement)(from el in newHouse.Descendants("Components")
                                       select el).First();

            comp.Add(
                new XElement("Floor",
                new XAttribute("adjacentEnclosedSpace", "false"),
                new XAttribute("id", maxID),
                    new XElement("Label", name),
                    new XElement("Construction",
                        new XElement("Type", "User specified",
                        new XAttribute("rValue", floorRValue),
                        new XAttribute("nominalInsulation", "6.0582"))),
                    new XElement("Measurements",
                        new XAttribute("area", areaMetric),
                        new XAttribute("length", lengthMetric))));
            maxID++;
            return newHouse;
        }
        public XDocument ExtraFirstWalls()
        {
            string column = "H";
            int startRow = 22;
            int currentRow = startRow;
            string name;
            string corners;
            string intersections;
            string height;
            string perim;

            for (int i = startRow; i <= 23; i++)
            {
                string currentCell = column + currentRow.ToString();
                if ((GetCellValue("Calc", currentCell) != null) && double.Parse(GetCellValue("Calc", currentCell)) > 0)
                {
                    perim = GetCellValue("Calc", currentCell);
                    height = GetCellValue("Calc", "G" + currentRow);
                    name = GetCellValue("Calc", "A" + currentRow);
                    corners = GetCellValue("Calc", "E" + currentRow);
                    intersections = GetCellValue("Calc", "F" + currentRow);

                    NewWall(name, corners, intersections, height, perim);

                }
                currentRow++;
            }
            return newHouse;
        }

        public XDocument ExtraSecondWalls()
        {
            string column = "H";
            int startRow = 22;
            int endRow = 33;
            int currentRow = startRow;
            string name;
            string corners;
            string intersections;
            string height;
            string perim;

            for (int i = startRow; i <= endRow; i++)
            {
                string currentCell = column + currentRow.ToString();
                if ((GetCellValue("Calc", currentCell) != null) && double.Parse(GetCellValue("Calc", currentCell)) > 0)
                {
                    perim = GetCellValue("Calc", currentCell);
                    height = GetCellValue("Calc", "G" + currentRow);
                    name = GetCellValue("Calc", "A" + currentRow);
                    corners = GetCellValue("Calc", "E" + currentRow);
                    intersections = GetCellValue("Calc", "F" + currentRow);

                    NewWall(name, corners, intersections, height, perim);

                }
                currentRow++;
                if (currentRow == 24)
                {
                    currentRow++;
                    i++;
                }
                if (currentRow == 30)
                {
                    currentRow += 2;
                    i += 2;
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

    }
}
