import { CLIENT_MATRIX_CATEGORY_COLUMNS } from "@/lib/services/bulkUpload/clientTemplateHeaders";

/**
 * Official NPORS Nov 2023 titles (from NPORS-Category-List-Nov23.pdf),
 * shortened for admin UI. Codes must stay as N### for SharePoint / matrix sync.
 *
 * Where the Training Matrix template already has a short name, that wins.
 */
const OFFICIAL_SHORT_TITLES: Record<string, string> = {
  N001: "Ind FLT",
  N003: "Reach Lift Truck",
  N004: "Lorry Mounted Lift Truck",
  N005: "Order Picker",
  N006: "Side Loader",
  N007: "Narrow Aisle Lift Truck",
  N008: "Pallet / Stacker Truck",
  N009: "Rough Terrain Lift Truck",
  N010: "Telescopic Handler",
  N012: "Container Handler",
  N013: "Pivot Steer Lift Truck",
  N014: "Marine Knuckle Boom Crane",
  N015: "Soil Displacement Hammer",
  N016: "Micro Excavator 360° ≤1t",
  N017: "Abrasive Wheels – Cut-off Saw",
  N019: "Thorough Examination Report",
  N020: "Tiltrotator System",
  N021: "Suction Excavator",
  N022: "Piling Rig Attendant",
  N025: "Concrete Cutting Chainsaw",
  N026: "Roll Crusher – Packer",
  N027: "Excavation Marshal",
  N029: "Refuse Collection Vehicle",
  N030: "Machine Control Systems",
  N031: "Asbestos Awareness",
  N033: "Puller Winch",
  N034: "Plant Supervisor Awareness",
  N035: "MEWP Supervisor Awareness",
  N036: "FLT Supervisor Awareness",
  N038: "High Pressure Water Jetting",
  N039: "Plant Inspector – Thorough Exam",
  N040: "Soil Stabiliser",
  N041: "Post Rammer",
  N045: "Mobile Loading Ramps",
  N046: "Rigging Planner",
  N047: "Rigging",
  N048: "Powered Handheld Breaker",
  N049: "Shoring",
  N050: "3D GPS Machine Control",
  N100: "Excavator as a Crane",
  N101: "Mobile Crane",
  N102: "Tower Crane",
  N103: "Crawler Crane",
  N104: "Overhead Crane",
  N105: "Quay Crane",
  N106: "Log Handler",
  N107: "Lorry Loader",
  N108: "MEWP Boom",
  N109: "MEWP Scissor",
  N111: "Hoist",
  N114: "Overhead Container Gantry Crane",
  N115: "Remote Control Tower Crane",
  N116: "Dragline",
  N119: "Mast Climber",
  N120: "Plant Loader & Securer",
  N130: "Horizontal Directional Drilling",
  N132: "Plant Mover",
  N133: "Plant Machinery Marshal",
  N135: "Multi Directional Lift Truck",
  N136: "Boat Hoist",
  N138: "Telehandler Suspended Loads",
  N139: "Forward Tipping Mini Dumper ≤1t",
  N140: "Compact Crane",
  N201: "Excavator 180°",
  N202: "Excavator 360°",
  N203: "Trencher",
  N204: "Forward Tipping Dumper",
  N205: "Rear Tipping Dumper",
  N206: "Loader Compressor",
  N207: "Crusher",
  N208: "Screener",
  N209: "Loading Shovel",
  N210: "Tow Tractor",
  N211: "Concrete Pump (Mobile)",
  N212: "Skidsteer Loader",
  N214: "Road Roller",
  N215: "Dozer",
  N216: "Road Planer",
  N217: "Road Sweeper",
  N219: "Skip Loader",
  N220: "Paver",
  N221: "Piling Rig",
  N223: "Scraper",
  N224: "Landfill Compactor",
  N225: "Hook Loader Vehicle",
  N240: "Grader",
  N241: "Pay Welder",
  N242: "Side Boom",
  N243: "Shunter Vehicle",
  N244: "Chipper",
  N301: "Abrasive Wheels – Practical",
  N301A: "Abrasive Wheels Awareness",
  N302: "Bench Saw",
  N303: "Circular Saw",
  N304: "Cable Avoidance Tool",
  N401: "Appointed Person",
  N402: "Slinger / Signaller",
  N403: "Vehicle Marshal",
  N404: "Safe Working at Height",
  N405: "Crane / Lift Supervisor",
  N601: "Agricultural Tractor",
  N602: "Chainsaw – Cross Cutting",
  N603: "Wood Chipper / Shredder",
  N604: "Grass Cutters / Mowers",
  N605: "Strimmer / Brushcutter",
  N606: "Stump Grinder",
  N607: "4x4 Off Road Vehicle",
  N608: "All-Terrain Vehicle",
  N609: "Winching & Recovery",
  N610: "Hand Held Hedge Trimmer",
  N702A: "Confined Spaces – Low Risk",
  N702B: "Confined Spaces – Medium Risk",
  N702C: "Confined Spaces – High Risk",
  N703: "Fire Warden",
  N704: "Manual Handling",
  N706: "Mobile Towers",
  N710: "Cartridge Tools",
  N711: "Safe Use of Ladders",
  N714: "Safety at Street & Road Works",
  N721: "Static Concrete Placing Boom",
  N722: "Materials Re-Handler",
  N723: "Harness and Fall Arrest",
  N726: "Quick Hitch Awareness",
  N802: "Gritter / Snowplough",
};

/**
 * Admin NPORS category options — matrix columns with short official titles.
 * Values are category codes (must match SharePoint MultiChoice where provisioned).
 */
export const NPORS_CATEGORY_CHOICES = CLIENT_MATRIX_CATEGORY_COLUMNS.map(
  (column) => {
    const code = column.code.toUpperCase();
    return {
      code,
      title: OFFICIAL_SHORT_TITLES[code] ?? column.name,
    };
  },
);

/**
 * Admin NPORS category dropdown — labels show "CODE – Short title".
 */
export function getNporsCategoryOptions(): Array<{
  value: string;
  label: string;
}> {
  return NPORS_CATEGORY_CHOICES.map(({ code, title }) => ({
    value: code,
    label: `${code} – ${title}`,
  }));
}

/** Lookup a short display title for an NPORS code. */
export function getNporsShortTitle(code: string | null | undefined): string | null {
  if (!code?.trim()) return null;
  const key = code.trim().toUpperCase();
  return OFFICIAL_SHORT_TITLES[key] ?? null;
}
