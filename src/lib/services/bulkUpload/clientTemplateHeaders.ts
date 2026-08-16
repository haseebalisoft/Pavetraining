/** Auto-aligned to client Excel templates (Company list.xlsx + Workforce list.xlsx + Training matrix example.xlsx). */
export const CLIENT_COMPANY_HEADERS = [
  "Company Number",
  "Company Name",
  "Company Size",
  "Registered Address",
  "Company Reg Number",
  "VAT No",
  "Tel No",
  "Email",
  "Main Contact",
  "Accounts Contact Name",
  "Accounts address",
  "Accounts Contact number",
  "Accounts email",
  "Notes prices agreed",
  "Company Logo",
  "Status",
] as const;

export const CLIENT_WORKFORCE_HEADERS = [
  "Workforce Number",
  "Candidate Name",
  "Company Name",
  "Company Number",
  "Training manager",
  "Training manager email",
  "Supervisor",
  "Supervisor email",
  "Candidate Address",
  "Email",
  "Contact number",
  "Date of birth",
  "Ni Number",
  "NPORS Number",
  "CSCS Number",
  "Cscs Expiry",
  "SWQR Number",
  "Swqr Expiry",
  "EUSR Number",
  "Eusr Expiry",
  "In House Certification Number",
  "Department",
  "Candidate Photo",
  "Status",
  "Notes"
] as const;

export type MatrixCategoryColumn = {
  header: string;
  code: string;
  name: string;
  /** When set, also written to Training Matrix list expiry column. */
  matrixField: string | null;
};

export const CLIENT_MATRIX_META_HEADERS = [
  "Name",
  "DOB",
  "CSCS Expiry",
  "SSSTS Expiry",
  "SMSTS Expiry",
  "NRSWA Expiry",
  "EUSR Expiry",
  /** Present in Training matrix example.xlsx; omitted in SharePoint list export 11-07-26. */
  "Face ift",
] as const;

export const CLIENT_MATRIX_CATEGORY_COLUMNS: MatrixCategoryColumn[] = [
  {
    "header": "N001 - Ind FLT",
    "code": "N001",
    "name": "Ind FLT",
    "matrixField": "n001Expiry"
  },
  {
    "header": "N003 - Reach Lift Truck",
    "code": "N003",
    "name": "Reach Lift Truck",
    "matrixField": "n003Expiry"
  },
  {
    "header": "N004 - Lorry Mounted Lift Truck",
    "code": "N004",
    "name": "Lorry Mounted Lift Truck",
    "matrixField": "n004Expiry"
  },
  {
    "header": "N006 - Side Loader",
    "code": "N006",
    "name": "Side Loader",
    "matrixField": null
  },
  {
    "header": "N009 - Rough Terrain Lift Truck",
    "code": "N009",
    "name": "Rough Terrain Lift Truck",
    "matrixField": null
  },
  {
    "header": "N010 - Telescopic Handler",
    "code": "N010",
    "name": "Telescopic Handler",
    "matrixField": "n010Expiry"
  },
  {
    "header": "N135 - Multi Di Lift Truck",
    "code": "N135",
    "name": "Multi Di Lift Truck",
    "matrixField": null
  },
  {
    "header": "N108 - MEWP Boom",
    "code": "N108",
    "name": "MEWP Boom",
    "matrixField": null
  },
  {
    "header": "N109 - MEWP Scissor",
    "code": "N109",
    "name": "MEWP Scissor",
    "matrixField": null
  },
  {
    "header": "N111 - Hoist",
    "code": "N111",
    "name": "Hoist",
    "matrixField": null
  },
  {
    "header": "N119 - Mast Climber",
    "code": "N119",
    "name": "Mast Climber",
    "matrixField": null
  },
  {
    "header": "N706 - Mobile Towers",
    "code": "N706",
    "name": "Mobile Towers",
    "matrixField": null
  },
  {
    "header": "N014 - Marine Knuckle Boom Crane",
    "code": "N014",
    "name": "Marine Knuckle Boom Crane",
    "matrixField": null
  },
  {
    "header": "N019 - Report of  Thorough Exam",
    "code": "N019",
    "name": "Report of  Thorough Exam",
    "matrixField": null
  },
  {
    "header": "N039 - Plant Inspector – Thorough Exam",
    "code": "N039",
    "name": "Plant Inspector – Thorough Exam",
    "matrixField": null
  },
  {
    "header": "N046 - Rigging Planner",
    "code": "N046",
    "name": "Rigging Planner",
    "matrixField": null
  },
  {
    "header": "N047 - Rigging",
    "code": "N047",
    "name": "Rigging",
    "matrixField": null
  },
  {
    "header": "N100 - Exc Crane",
    "code": "N100",
    "name": "Exc Crane",
    "matrixField": "n100Expiry"
  },
  {
    "header": "N101 - Mobile Crane",
    "code": "N101",
    "name": "Mobile Crane",
    "matrixField": null
  },
  {
    "header": "N102 - Tower Crane",
    "code": "N102",
    "name": "Tower Crane",
    "matrixField": null
  },
  {
    "header": "N103 - Crawler Crane",
    "code": "N103",
    "name": "Crawler Crane",
    "matrixField": null
  },
  {
    "header": "N104 - Overhead Crane",
    "code": "N104",
    "name": "Overhead Crane",
    "matrixField": null
  },
  {
    "header": "N105 - Quay Crane",
    "code": "N105",
    "name": "Quay Crane",
    "matrixField": null
  },
  {
    "header": "N106 - Log Handler",
    "code": "N106",
    "name": "Log Handler",
    "matrixField": null
  },
  {
    "header": "N107 - Lorry Loader",
    "code": "N107",
    "name": "Lorry Loader",
    "matrixField": null
  },
  {
    "header": "N114 - Overhead Container Gantry Crane",
    "code": "N114",
    "name": "Overhead Container Gantry Crane",
    "matrixField": null
  },
  {
    "header": "N115 - Remote Control Tower Crane",
    "code": "N115",
    "name": "Remote Control Tower Crane",
    "matrixField": null
  },
  {
    "header": "N138 - Tele Suspended Loads",
    "code": "N138",
    "name": "Tele Suspended Loads",
    "matrixField": null
  },
  {
    "header": "N140 - Compact Crane",
    "code": "N140",
    "name": "Compact Crane",
    "matrixField": null
  },
  {
    "header": "N242 - Side Boom",
    "code": "N242",
    "name": "Side Boom",
    "matrixField": null
  },
  {
    "header": "N401 - Appointed Person",
    "code": "N401",
    "name": "Appointed Person",
    "matrixField": null
  },
  {
    "header": "N402 - Slinger / Signaller",
    "code": "N402",
    "name": "Slinger / Signaller",
    "matrixField": null
  },
  {
    "header": "N405 - Crane / Lift Supervisor",
    "code": "N405",
    "name": "Crane / Lift Supervisor",
    "matrixField": null
  },
  {
    "header": "N722 - Materials Re-Handler",
    "code": "N722",
    "name": "Materials Re-Handler",
    "matrixField": null
  },
  {
    "header": "N016 - Micro Excavator 360° up to 1 Tonne",
    "code": "N016",
    "name": "Micro Excavator 360° up to 1 Tonne",
    "matrixField": null
  },
  {
    "header": "N020 - Tiltrotator System",
    "code": "N020",
    "name": "Tiltrotator System",
    "matrixField": "n020Expiry"
  },
  {
    "header": "N021 - Suction Excavator",
    "code": "N021",
    "name": "Suction Excavator",
    "matrixField": "n021Expiry"
  },
  {
    "header": "N030 - Machine Control Systems",
    "code": "N030",
    "name": "Machine Control Systems",
    "matrixField": null
  },
  {
    "header": "N116 - Dragline",
    "code": "N116",
    "name": "Dragline",
    "matrixField": null
  },
  {
    "header": "N139 - Forward Tipping Mini Dumper up to 1 Tonne",
    "code": "N139",
    "name": "Forward Tipping Mini Dumper up to 1 Tonne",
    "matrixField": null
  },
  {
    "header": "N201 - Excavator 180°",
    "code": "N201",
    "name": "Excavator 180°",
    "matrixField": null
  },
  {
    "header": "N202 - Excavator 360°",
    "code": "N202",
    "name": "Excavator 360°",
    "matrixField": null
  },
  {
    "header": "N203 - Trencher",
    "code": "N203",
    "name": "Trencher",
    "matrixField": null
  },
  {
    "header": "N204 - Forward Tipping Dumper",
    "code": "N204",
    "name": "Forward Tipping Dumper",
    "matrixField": null
  },
  {
    "header": "N205 - Rear Tipping Dumper",
    "code": "N205",
    "name": "Rear Tipping Dumper",
    "matrixField": null
  },
  {
    "header": "N206 - Loader Compressor",
    "code": "N206",
    "name": "Loader Compressor",
    "matrixField": null
  },
  {
    "header": "N209 - Loading Shovel",
    "code": "N209",
    "name": "Loading Shovel",
    "matrixField": null
  },
  {
    "header": "N212 - Skidsteer Loader",
    "code": "N212",
    "name": "Skidsteer Loader",
    "matrixField": null
  },
  {
    "header": "N215 - Dozer",
    "code": "N215",
    "name": "Dozer",
    "matrixField": null
  },
  {
    "header": "N223 - Scraper",
    "code": "N223",
    "name": "Scraper",
    "matrixField": null
  },
  {
    "header": "N224 - Landfill Compactor",
    "code": "N224",
    "name": "Landfill Compactor",
    "matrixField": null
  },
  {
    "header": "N240 - Grader",
    "code": "N240",
    "name": "Grader",
    "matrixField": null
  },
  {
    "header": "N241 - Pay Welder",
    "code": "N241",
    "name": "Pay Welder",
    "matrixField": null
  },
  {
    "header": "N601 - Agricultural Tractor",
    "code": "N601",
    "name": "Agricultural Tractor",
    "matrixField": null
  },
  {
    "header": "N726 - Quick Hitch Awareness",
    "code": "N726",
    "name": "Quick Hitch Awareness",
    "matrixField": null
  },
  {
    "header": "N214 - Road Roller",
    "code": "N214",
    "name": "Road Roller",
    "matrixField": null
  },
  {
    "header": "N216 - Road Planer",
    "code": "N216",
    "name": "Road Planer",
    "matrixField": null
  },
  {
    "header": "N217 - Road Sweeper",
    "code": "N217",
    "name": "Road Sweeper",
    "matrixField": null
  },
  {
    "header": "N220 - Paver",
    "code": "N220",
    "name": "Paver",
    "matrixField": null
  },
  {
    "header": "N244 - Chipper",
    "code": "N244",
    "name": "Chipper",
    "matrixField": null
  },
  {
    "header": "N609 - Winching & Recovery",
    "code": "N609",
    "name": "Winching & Recovery",
    "matrixField": null
  },
  {
    "header": "N802 - Gritter / Snowplough",
    "code": "N802",
    "name": "Gritter / Snowplough",
    "matrixField": null
  },
  {
    "header": "N027 - Excavation Marshal - Banksperson",
    "code": "N027",
    "name": "Excavation Marshal - Banksperson",
    "matrixField": "n027Expiry"
  },
  {
    "header": "N029 - Refuse Collection Vehicle",
    "code": "N029",
    "name": "Refuse Collection Vehicle",
    "matrixField": null
  },
  {
    "header": "N120 - Plant Loader & Securer",
    "code": "N120",
    "name": "Plant Loader & Securer",
    "matrixField": null
  },
  {
    "header": "N210 - Tow Tractor",
    "code": "N210",
    "name": "Tow Tractor",
    "matrixField": null
  },
  {
    "header": "N219 - Skip Loader",
    "code": "N219",
    "name": "Skip Loader",
    "matrixField": null
  },
  {
    "header": "N225 - Multi Lift & Drop – Hook Loader Vehicle",
    "code": "N225",
    "name": "Multi Lift & Drop – Hook Loader Vehicle",
    "matrixField": null
  },
  {
    "header": "N243 - Shunter Vehicle",
    "code": "N243",
    "name": "Shunter Vehicle",
    "matrixField": null
  },
  {
    "header": "N607 - 4 x 4 Off Road Vehicle",
    "code": "N607",
    "name": "4 x 4 Off Road Vehicle",
    "matrixField": null
  },
  {
    "header": "N608 - All-Terrain Vehicle",
    "code": "N608",
    "name": "All-Terrain Vehicle",
    "matrixField": null
  },
  {
    "header": "N045 - Safe Use of Mobile Loading Ramps",
    "code": "N045",
    "name": "Safe Use of Mobile Loading Ramps",
    "matrixField": null
  },
  {
    "header": "N017 - Abrasive Wheels – Hand Held Cut off Saw",
    "code": "N017",
    "name": "Abrasive Wheels – Hand Held Cut off Saw",
    "matrixField": null
  },
  {
    "header": "N025 - Concrete Cutting Chainsaw",
    "code": "N025",
    "name": "Concrete Cutting Chainsaw",
    "matrixField": null
  },
  {
    "header": "N301A - Abrasive Wheels Awareness",
    "code": "N301A",
    "name": "Abrasive Wheels Awareness",
    "matrixField": null
  },
  {
    "header": "N301 - Abrasive Wheels with Practical Cutting / Grinding",
    "code": "N301",
    "name": "Abrasive Wheels with Practical Cutting / Grinding",
    "matrixField": null
  },
  {
    "header": "N302 - Bench Saw",
    "code": "N302",
    "name": "Bench Saw",
    "matrixField": null
  },
  {
    "header": "N303 - Circular Saw",
    "code": "N303",
    "name": "Circular Saw",
    "matrixField": null
  },
  {
    "header": "N304 - Cable Avoidance Tool",
    "code": "N304",
    "name": "Cable Avoidance Tool",
    "matrixField": null
  },
  {
    "header": "N602 - Chainsaw – Maintenance and Cross Cutting",
    "code": "N602",
    "name": "Chainsaw – Maintenance and Cross Cutting",
    "matrixField": null
  },
  {
    "header": "N603 - Wood Chipper / Shredder",
    "code": "N603",
    "name": "Wood Chipper / Shredder",
    "matrixField": null
  },
  {
    "header": "N604 - Grass Cutters / Mowers",
    "code": "N604",
    "name": "Grass Cutters / Mowers",
    "matrixField": null
  },
  {
    "header": "N605 - Strimmer / Brushcutter",
    "code": "N605",
    "name": "Strimmer / Brushcutter",
    "matrixField": null
  },
  {
    "header": "N606 - Stump Grinder",
    "code": "N606",
    "name": "Stump Grinder",
    "matrixField": null
  },
  {
    "header": "N610 - Hand Held Hedge Trimmer",
    "code": "N610",
    "name": "Hand Held Hedge Trimmer",
    "matrixField": null
  },
  {
    "header": "N710 - Cartridge Tools",
    "code": "N710",
    "name": "Cartridge Tools",
    "matrixField": null
  },
  {
    "header": "N048 - Powered Handheld Breaker",
    "code": "N048",
    "name": "Powered Handheld Breaker",
    "matrixField": null
  },
  {
    "header": "N033 - Puller Winch",
    "code": "N033",
    "name": "Puller Winch",
    "matrixField": null
  },
  {
    "header": "N038 - High Pressure Water Jetting",
    "code": "N038",
    "name": "High Pressure Water Jetting",
    "matrixField": null
  },
  {
    "header": "N049 - Shoring (Install, Inspect, Remove)",
    "code": "N049",
    "name": "Shoring (Install, Inspect, Remove)",
    "matrixField": null
  },
  {
    "header": "N050 - 3D GPS Machine Control System",
    "code": "N050",
    "name": "3D GPS Machine Control System",
    "matrixField": null
  },
  {
    "header": "N015 - Soil Displacement Hammer",
    "code": "N015",
    "name": "Soil Displacement Hammer",
    "matrixField": null
  },
  {
    "header": "N022 - Piling Rig Attendant",
    "code": "N022",
    "name": "Piling Rig Attendant",
    "matrixField": null
  },
  {
    "header": "N026 - Roll Crusher – Packer",
    "code": "N026",
    "name": "Roll Crusher – Packer",
    "matrixField": null
  },
  {
    "header": "N130 - Horizontal Directional Drilling Rig",
    "code": "N130",
    "name": "Horizontal Directional Drilling Rig",
    "matrixField": null
  },
  {
    "header": "N207 - Crusher",
    "code": "N207",
    "name": "Crusher",
    "matrixField": null
  },
  {
    "header": "N208 - Screener",
    "code": "N208",
    "name": "Screener",
    "matrixField": null
  },
  {
    "header": "N211 - Concrete Pump (Mobile)",
    "code": "N211",
    "name": "Concrete Pump (Mobile)",
    "matrixField": null
  },
  {
    "header": "N221 - Piling Rig",
    "code": "N221",
    "name": "Piling Rig",
    "matrixField": null
  },
  {
    "header": "N721 - Static Concrete Placing Boom",
    "code": "N721",
    "name": "Static Concrete Placing Boom",
    "matrixField": null
  },
  {
    "header": "N040 - Soil Stabiliser",
    "code": "N040",
    "name": "Soil Stabiliser",
    "matrixField": null
  },
  {
    "header": "N041 - Post Rammer",
    "code": "N041",
    "name": "Post Rammer",
    "matrixField": null
  },
  {
    "header": "N031 - Asbestos Awareness",
    "code": "N031",
    "name": "Asbestos Awareness",
    "matrixField": "n031Expiry"
  },
  {
    "header": "N034 - Plant Supervisor Awareness",
    "code": "N034",
    "name": "Plant Supervisor Awareness",
    "matrixField": null
  },
  {
    "header": "N035 - MEWP Supervisor Awareness",
    "code": "N035",
    "name": "MEWP Supervisor Awareness",
    "matrixField": null
  },
  {
    "header": "N036 - FLT Supervisor Awareness",
    "code": "N036",
    "name": "FLT Supervisor Awareness",
    "matrixField": null
  },
  {
    "header": "N132 - Plant Mover – Non Operational Duties",
    "code": "N132",
    "name": "Plant Mover – Non Operational Duties",
    "matrixField": null
  },
  {
    "header": "N133 - Plant Machinery Marshal",
    "code": "N133",
    "name": "Plant Machinery Marshal",
    "matrixField": null
  },
  {
    "header": "N403 - Vehicle Marshal",
    "code": "N403",
    "name": "Vehicle Marshal",
    "matrixField": null
  },
  {
    "header": "N404 - Safe Working at Height",
    "code": "N404",
    "name": "Safe Working at Height",
    "matrixField": null
  },
  {
    "header": "N702A - Confined Spaces – Low Risk",
    "code": "N702A",
    "name": "Confined Spaces – Low Risk",
    "matrixField": null
  },
  {
    "header": "N702B - Confined Spaces – Medium Risk",
    "code": "N702B",
    "name": "Confined Spaces – Medium Risk",
    "matrixField": null
  },
  {
    "header": "N702C - Confined Spaces – High Risk",
    "code": "N702C",
    "name": "Confined Spaces – High Risk",
    "matrixField": null
  },
  {
    "header": "N703 - Fire Warden",
    "code": "N703",
    "name": "Fire Warden",
    "matrixField": null
  },
  {
    "header": "N704 - Manual Handling",
    "code": "N704",
    "name": "Manual Handling",
    "matrixField": null
  },
  {
    "header": "N711 - Safe Use of Ladders and Ladders",
    "code": "N711",
    "name": "Safe Use of Ladders and Ladders",
    "matrixField": null
  },
  {
    "header": "N714 - Safety at Street and Road Works",
    "code": "N714",
    "name": "Safety at Street and Road Works",
    "matrixField": null
  },
  {
    "header": "N723 - Harness and Fall Arrest",
    "code": "N723",
    "name": "Harness and Fall Arrest",
    "matrixField": null
  }
];

/** Exact column order from Training matrix example.xlsx (for admin table + import). */
export const CLIENT_MATRIX_DISPLAY_HEADERS: string[] = [
  ...CLIENT_MATRIX_META_HEADERS,
  ...CLIENT_MATRIX_CATEGORY_COLUMNS.map((column) => column.header),
];
