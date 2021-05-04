//Created by Jesse Russo 2019
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace HotREF
{
    class CreateRef
    {
        string ceilingRValue = "10.4292";
        string wallRValue = "3.0802";
        string garWallRValue = "2.9199";
        string bsmtWallRValue = "3.3443";
        string floorRValue = "5.0191";
        string garFloorRValue = "4.8589";
        string furnaceEF = "92";
        string slabRValue = "1.902";
        string windowRValue = "0.6252";
        string doorRValue = "0.6252";
        string weatherZone = "7A";
        int maxID;
        int codeID = 3;

        public CreateRef(XDocument house, string zone)
        {
            List<char> codeIDs = new List<char>();
            var hasCode = from el in house.Descendants("Codes").Descendants().Attributes("id")
                          select el;
            foreach(string code in hasCode)
            {
                codeIDs.Add((code.Last()));
            }
            weatherZone = zone;
            SetZone();
        }

        private void SetZone()
        {
            if (weatherZone.Equals("Zone 6"))
            {
                ceilingRValue = "8.6699";
                floorRValue = "4.6704";
                bsmtWallRValue = "2.8636";
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
                        ceiling.Element("Construction").Element("CeilingType").SetAttributeValue("rValue", ceilingRValue);
                    }
                }
            }
            //Changes R values for walls and rim joists
            foreach (XElement wall in house.Descendants("Wall"))
            {
                //Check for garage walls
                foreach (XElement type in wall.Descendants("Type"))
                    if (wall.Element("Label").Value.ToString().Contains("Garage"))
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
            foreach (XElement floor in house.Descendants("Floor"))
            {
                double area = System.Convert.ToDouble(floor.Element("Measurements").Attribute("area").Value.ToString());
                foreach (XElement type in floor.Descendants("Type"))
                {
                    //Check if floor is over 107sq ft. If so set as ExpGar R value
                    if (area > 10)
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
        public XDocument HvacChanger(XDocument house, string furnaceOutput)
        {
            //Convert BTUs/h entered to KW
            double btus = System.Convert.ToDouble(furnaceOutput);
            btus = Math.Round((btus * 0.00029307107), 5);
            furnaceOutput = btus.ToString();

            // Changes furnace output capacity and EF values
            foreach (XElement furnace in house.Descendants("Furnace").Descendants("Specifications"))
            {
                    furnace.SetAttributeValue("efficiency", furnaceEF);
                    furnace.Element("OutputCapacity").SetAttributeValue("value", furnaceOutput);   
            }
            //Changes blower door test value to 2.5
            foreach(XElement bt in house.Descendants("BlowerTest"))
            {
                bt.SetAttributeValue("airChangeRate", "2.5");
            }
            //Changes A/C SEER
            foreach (XElement ac in house.Descendants("AirConditioning").Descendants("Efficiency"))
            {
                ac.SetAttributeValue("value", "14.5");
            }
            return house;
        }

        public XDocument HotWater(XDocument house, string tankEF)
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
                    tank.Element("EnergyFactor").SetAttributeValue("value", tankEF);
                }
                else
                {
                    foreach (XElement ef in tank.Descendants("EnergyFactor"))
                    {
                        ef.SetAttributeValue("value", tankEF);
                    }
                }
            }            
            return house;
        }

        //Adds utility ventilation element and fills values for vent rate and fan power
        public XDocument AddHrv(XDocument house, string ventilation, string fanPower)
        {
            foreach (XElement vent in house.Descendants("WholeHouseVentilatorList"))
            {
                double ls = System.Convert.ToDouble(ventilation);
                ls = Math.Round((ls * 0.47195), 4);
                ventilation = ls.ToString();

                vent.Add(
                    new XElement("BaseVentilator",
                    new XAttribute("supplyFlowrate", ventilation),
                    new XAttribute("exhaustFlowrate", ventilation),
                    new XAttribute("fanPower1", fanPower),
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

        public XDocument Doors(XDocument house, string doorSize)
        {
            double width = Math.Round((System.Convert.ToDouble(doorSize) * 0.0254), 4);
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

        public XDocument Windows(XDocument house, string windowSize)
        {
            double size = Math.Round((System.Convert.ToDouble(windowSize) * 25.4), 6);
            string floors = "2nd Flr";
            List<string> wallList = new List<string>();

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
                            //Add North Window
                            comp.Add(
                                new XElement("Window",
                                new XAttribute("number", "1"),
                                new XAttribute("er", "-32.1684"),
                                new XAttribute("shgc", "0.26"),
                                new XAttribute("adjacentEnclosedSpace", "false"),
                                new XAttribute("id", maxID),
                                    new XElement("Label", "North"),
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
                                    new XAttribute("code", "5"),
                                    new XElement("English", "North"),
                                    new XElement("French", "Nord"))));
                            maxID++;
                            //Add South window
                            comp.Add(
                                new XElement("Window",
                                new XAttribute("number", "1"),
                                new XAttribute("er", "-32.1684"),
                                new XAttribute("shgc", "0.26"),
                                new XAttribute("adjacentEnclosedSpace", "false"),
                                new XAttribute("id", maxID),
                                    new XElement("Label", "South"),
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
                                    new XAttribute("code", "1"),
                                    new XElement("English", "South"),
                                    new XElement("French", "Sud"))));
                            maxID++;
                            //Add East window
                            comp.Add(
                                new XElement("Window",
                                new XAttribute("number", "1"),
                                new XAttribute("er", "-32.1684"),
                                new XAttribute("shgc", "0.26"),
                                new XAttribute("adjacentEnclosedSpace", "false"),
                                new XAttribute("id", maxID),
                                    new XElement("Label", "East"),
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
                                    new XAttribute("code", "3"),
                                    new XElement("English", "East"),
                                    new XElement("French", "Est"))));
                            maxID++;
                            //Add West window
                            comp.Add(
                                new XElement("Window",
                                new XAttribute("number", "1"),
                                new XAttribute("er", "-32.1684"),
                                new XAttribute("shgc", "0.26"),
                                new XAttribute("adjacentEnclosedSpace", "false"),
                                new XAttribute("id", maxID),
                                    new XElement("Label", "West"),
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
                                    new XAttribute("code", "7"),
                                    new XElement("English", "West"),
                                    new XElement("French", "Ouest"))));
                            maxID++;
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
    }
}