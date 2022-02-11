//Created by Jesse Russo 2019
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;


namespace HotREF
{
    class CreateRef
    {
        private string CeilingRValue { get; set; } = "10.4292";
        private string wallRValue = "3.0802";
        private string garWallRValue = "2.9199";
        private string bsmtWallRValue = "3.3443";
        private string floorRValue = "5.0191";
        private string garFloorRValue = "4.8589";
        private string furnaceEF = "92";
        private string slabRValue = "1.902";
        private string windowRValue = "0.6252";
        private string doorRValue = "0.6252";
        private string weatherZone = "7A";
        private string acSeer = "14.5";
        private string heatedSlab = "15.8";
        private string ACH = "2.5";
        private int maxID;
        private int codeID = 3;
        private string ExcelFilePath {get; set;}

        public CreateRef(XDocument house, string zone, string excelPath)
        {
            List<char> codeIDs = new List<char>();
            var hasCode = from el in house.Descendants("Codes").Descendants().Attributes("id")
                          select el.Value;

            foreach(string code in hasCode)
            {
                codeIDs.Add((code.Last()));
            }
            weatherZone = zone;
            SetZone();
            ExcelFilePath = excelPath;
        }

        private void SetZone()
        {
            if (weatherZone.Equals("Zone 6"))
            {
                CeilingRValue = "8.6699";
                floorRValue = "4.6704";
                bsmtWallRValue = "2.8636";
                heatedSlab = "2.78254";
            }
        }

        //Finds the highest value ID attribute in order for new windows and doors to have 
        //their IDs incremented from maxID
        public void FindID(XDocument house)
        {
            List<string> ids = new List<string>();
            var hasID =
                from el in house.Descendants("House").Descendants().Attributes("id")
                where el.Value != null
                select el.Value;
            foreach (string id in hasID)
            {
                ids.Add(id);
            }

            List<int> idList = ids.Select(s => int.Parse(s)).ToList();
            maxID = idList.Max() + 1;
        }
        public XDocument Remover(XDocument house)
        {
            house.Descendants("Door").Remove();
            house.Descendants("Window").Remove();
            house.Descendants("Hrv").Remove();
            house.Descendants("BaseVentilator").Remove();
            return house;
        }
        //Changes R value for ceilings, walls, 
        public XDocument RChanger(XDocument house)
        {
            //Changes R values of ceiling elements
            foreach (XElement ceiling in house.Descendants("Ceiling"))
            {
                string flat = "Flat";
                string cath = "Cathedral";
                foreach (XElement type in ceiling.Descendants("CeilingType"))
                {
                    if (cath.Equals(ceiling.Element("Construction").Element("Type").Element("English").Value.ToString()))
                    {
                        ceiling.Element("Construction").Element("CeilingType").SetAttributeValue("rValue", floorRValue);
                    }
                    else if (flat.Equals(ceiling.Element("Construction").Element("Type").Element("English").Value.ToString()))
                    {
                        ceiling.Element("Construction").Element("CeilingType").SetAttributeValue("rValue", floorRValue);
                    }
                    else
                    {
                        ceiling.Element("Construction").Element("CeilingType").SetAttributeValue("rValue", CeilingRValue);
                    }
                }
            }
            //Changes R values for walls and rim joists
            foreach (XElement wall in house.Descendants("Wall"))
            {
                //Check for garage walls
                foreach (XElement type in wall.Descendants("Type"))
                    if (wall.Element("Label").Value.ToString().ToLower().Contains("garage"))
                    {
                        //If wall is a garage wall
                        type.SetAttributeValue("rValue", garWallRValue);
                    }
                    else
                    {
                        //If wall is NOT a garage wall
                        type.SetAttributeValue("rValue", wallRValue);
                    }
            }

            //Changes Basement wall insulation and pony wall R values
            foreach (XElement bsmt in house.Descendants("Basement"))
            {
                //Set R value for pony walls if they exist in bsmt
                foreach (XElement pWall in bsmt.Descendants("PonyWallType").Descendants("Section"))
                {
                        pWall.SetAttributeValue("rsi", wallRValue);
                }
                //Set R value for interior wall insulation for bsmt walls
                foreach (XElement intIns in bsmt.Descendants("InteriorAddedInsulation").Descendants("Section"))
                {
                        intIns.SetAttributeValue("rsi", bsmtWallRValue);
                }
                //Set R value for insulation added to slab
                foreach(XElement slab in bsmt.Descendants("AddedToSlab"))
                {
                    slab.SetAttributeValue("rValue", slabRValue);
                }
                //Set R value for bsmt floor headers
                foreach (XElement fh in bsmt.Descendants("FloorHeader").Descendants("Type"))
                {
                        fh.SetAttributeValue("rValue", wallRValue);
                }
            }

            foreach (XElement slab in house.Descendants("Slab").Descendants("Wall").Descendants("RValues"))
            {
                    slab.SetAttributeValue("thermalBreak", slabRValue); 
            }
            foreach (XElement slabIns in house.Descendants("Slab").Descendants("AddedToSlab"))
            {
                slabIns.SetAttributeValue("rValue", slabRValue);
            }

            // Set R value for exposed floors
            string GarFloorName = GetCellValue("Calc", "L21");
            foreach (XElement floor in house.Descendants("Floor"))
            {
                foreach (XElement type in floor.Descendants("Type"))
                {
                    //Check for garage floor
                    if (floor.Element("Label").Value == GarFloorName)
                    {
                        type.SetAttributeValue("rValue", garFloorRValue);
                    }
                    else
                    {
                        type.SetAttributeValue("rValue", floorRValue);
                    }
                }
            }
            return house;
        }
        //Changes values for furnace capacity, furnace EF, ACH, and DHW EF
        public XDocument HvacChanger(XDocument house)
        {
            //Convert BTUs/h to KW
            double btus = Convert.ToDouble(GetCellValue("General", "C6"));
            btus = Math.Round((btus * 0.00029307107), 5);

            // Changes furnace output capacity and EF values
            foreach (XElement furnace in house.Descendants("Furnace").Descendants("Specifications"))
            {
                    furnace.SetAttributeValue("efficiency", furnaceEF);
                    furnace.SetAttributeValue("isSteadyState", "false");
                    furnace.Element("OutputCapacity").SetAttributeValue("value", btus.ToString());   
            }
            //Changes blower door test value to 2.5
            foreach(XElement bt in house.Descendants("BlowerTest"))
            {
                bt.SetAttributeValue("airChangeRate", ACH);
            }
            //Changes A/C SEER
            foreach (XElement ac in house.Descendants("AirConditioning").Descendants("Efficiency"))
            {
                ac.SetAttributeValue("value", acSeer);
            }
            return house;
        }

        public XDocument HotWater(XDocument house)
        {
            //Changes DHW tank EF value
            foreach (XElement tank in house.Descendants("HotWater").Descendants("Primary"))
            {
                //check if instantaneous heater. Change to tank if true.
                if (tank.Element("TankType").Element("English").Value.ToString().Equals("Instantaneous (condensing)"))
                {
                    tank.SetAttributeValue("flueDiameter", "0");
                    tank.Element("TankType").SetAttributeValue("code", "9");
                    tank.Element("TankVolume").SetAttributeValue("code", "4");
                    tank.Element("EnergyFactor").SetAttributeValue("value", Math.Round(Convert.ToDouble(GetCellValue("General", "P5")),2));
                }
                else if(tank.Element("EnergySource").Attribute("code").Value == "1")
                {
                    tank.Element("EnergyFactor").SetAttributeValue("value", GetCellValue("General", "J31"));
                }
                else
                {
                    foreach (XElement ef in tank.Descendants("EnergyFactor"))
                    {
                        ef.SetAttributeValue("value", Math.Round(Convert.ToDouble(GetCellValue("General", "P5")), 2));
                    }
                }
            } 
            foreach(XElement tank in house.Descendants("HotWater").Descendants("Secondary"))
            {
                if (tank.Element("TankType").Element("English").Value.ToString().Equals("Instantaneous (condensing)"))
                {
                    tank.SetAttributeValue("flueDiameter", "0");
                    tank.Element("TankType").SetAttributeValue("code", "9");
                    tank.Element("TankVolume").SetAttributeValue("code", "4");
                    tank.Element("EnergyFactor").SetAttributeValue("value", Math.Round(Convert.ToDouble(GetCellValue("General", "P5")), 2));
                }
                if (tank.Element("EnergySource").Attribute("code").Value == "1")
                {
                    tank.Element("EnergyFactor").SetAttributeValue("value", GetCellValue("General", "J31"));
                }
                else
                {
                    tank.Element("EnergyFactor").SetAttributeValue("value", Math.Round(Convert.ToDouble(GetCellValue("General", "P5")),2));
                }
            }
            return house;
        }

        //Adds utility ventilation element and fills values for vent rate and fan power
        public XDocument AddFan(XDocument house)
        {
            foreach (XElement vent in house.Descendants("WholeHouseVentilatorList"))
            {
                double ls = Math.Round(Math.Round(Convert.ToDouble(GetCellValue("General", "H4")),1) * 0.47195,4);

                vent.Add(
                    new XElement("BaseVentilator",
                    new XAttribute("supplyFlowrate", ls.ToString()),
                    new XAttribute("exhaustFlowrate", ls.ToString()),
                    new XAttribute("fanPower1", Math.Round(Convert.ToDouble(GetCellValue("General", "K4")),1).ToString()),
                    new XAttribute("isDefaultFanpower", "false"),
                    new XAttribute("isEnergyStar", "false"),
                    new XAttribute("isHomeVentilatingInstituteCertified", "false"),
                    new XAttribute("isSupplemental", "false"),
                        new XElement("EquipmentInformation"),
                        new XElement("VentilatorType", 
                        new XAttribute("code", "4"),
                            new XElement("English", "Utility"),
                            new XElement("French", "Utilité"))));
            }
            return house;
        }

        public XDocument Doors(XDocument house)
        {
            double width = Math.Round((Convert.ToDouble(GetCellValue("General", "N10")) * 0.0254), 4);
            string ff = "1st Flr";
            foreach (XElement wall in house.Descendants("Wall"))
            {
                foreach (XElement comp in wall.Descendants("Components"))
                {
                    if (ff.Equals(wall.Element("Label").Value))
                    {
                        int i = 0;
                        while (i < 2)
                        {
                            comp.Add(
                                new XElement("Door",
                                new XAttribute("rValue", doorRValue),
                                new XAttribute("adjacentEnclosedSpace", "false"),
                                new XAttribute("id", maxID),
                                    new XElement("Label", "Door"),
                                    new XElement("Construction",
                                    new XAttribute("energyStar", "false"),
                                        new XElement("Type",
                                        new XAttribute("code", "8"),
                                        new XAttribute("value", doorRValue),
                                            new XElement("English", "User Specified"),
                                            new XElement("French", "Spécifié par l'utilisateur"))),
                                    new XElement("Measurements",
                                    new XAttribute("height", "2.1336"),
                                    new XAttribute("width", width))));
                            i++;
                            maxID++;
                        }
                    }
                }
            }
            return house;
        }

        public XDocument Windows(XDocument house)
        {
            double size = Math.Round((System.Convert.ToDouble(GetCellValue("General", "N9")) * 25.4), 6);
            string floors = "2nd Flr";
            List<string> wallList = new List<string>();
            Dictionary<string, string> facingDirection = new Dictionary<string, string>()
            {
                {"North", "5" },
                {"South", "1" },
                {"West", "7" },
                {"East", "3"},
            };

            var walls =
                from el in house.Descendants("House").Descendants("Wall").Descendants("Label")
                where el.Value != null
                select el.Value.ToString();

            foreach (string wall in walls)
            {
                wallList.Add(wall);
            }

            //Checks if a second floor exists. If not, windows are added to the first floor
            if (wallList.Contains("2nd Flr") != true)
            {
                floors = "1st Flr";
            }
            foreach (XElement wall in house.Descendants("Wall"))
            {
                foreach (XElement comp in wall.Descendants("Components"))
                {
                    if (floors.Equals(wall.Element("Label").Value))
                    {
                        foreach (KeyValuePair<string, string> pair in facingDirection)
                        {
                            comp.Add(
                                new XElement("Window",
                                new XAttribute("number", "1"),
                                new XAttribute("er", "-32.1684"),
                                new XAttribute("shgc", "0.26"),
                                new XAttribute("adjacentEnclosedSpace", "false"),
                                new XAttribute("id", maxID),
                                    new XElement("Label", pair.Key),
                                    new XElement("Construction",
                                    new XAttribute("energyStar", "false"),
                                        new XElement("Type", "ABCRef",
                                        new XAttribute("idref", ("Code " + codeID)),
                                        new XAttribute("rValue", windowRValue))),
                                    new XElement("Measurements",
                                    new XAttribute("height", size),
                                    new XAttribute("width", size),
                                    new XAttribute("headerHeight", "0"),
                                    new XAttribute("overhangWidth", "0"),
                                        new XElement("Tilt",
                                        new XAttribute("code", "1"),
                                        new XAttribute("value", "90"),
                                        new XElement("English", "Vertical"),
                                        new XElement("French", "Verticale"))),
                                    new XElement("Shading",
                                    new XAttribute("curtain", "1"),
                                    new XAttribute("shutterRValue", "0")),
                                    new XElement("FacingDirection",
                                    new XAttribute("code", pair.Value),
                                    new XElement("English", "North"),
                                    new XElement("French", "Nord"))));
                            maxID++;
                        }
                    }
                }
            }
            return house;
        }
        public XDocument AddCode(XDocument house)
            {
                var codes = from el in house.Descendants()
                            where el.Name == "Codes"
                            select el;

                foreach (XElement code in codes)
                {
                    code.AddFirst(
                        new XElement("Window",
                            new XElement("UserDefined",
                                new XElement("Code",
                                new XAttribute("id", ("Code " + codeID)),
                                new XAttribute("nominalRValue", "0"),
                                    new XElement("Label", "ABCRef"),
                                    new XElement("Description", ""),
                                    new XElement("Layers",
                                        new XElement("WindowLegacy",
                                        new XAttribute("frameHeight", "0"),
                                        new XAttribute("shgc", "0.26"),
                                        new XAttribute("rank", "1"),
                                            new XElement("Type",
                                            new XAttribute("code", 1),
                                                new XElement("English", "Picture"),
                                                new XElement("French", "Fixe")),
                                            new XElement("RsiValues",
                                            new XAttribute("centreOfGlass", windowRValue),
                                            new XAttribute("edgeOfGlass", windowRValue),
                                            new XAttribute("frame", windowRValue))))))));
                }
            return house;
        }
        //Method to get the value of single cells from worksheet
        public string GetCellValue(string sheetName, string refCell)
        {
            string value = null;

            using (FileStream fs = new FileStream(ExcelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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