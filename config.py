# =========================================================
# CONFIG.PY — MULTI-FLOOR INDUSTRIAL ENGINEERING REGISTRY
# Synchronized with machine_id.xlsx & 2026 NPT Cause Standard
# =========================================================

MAINTENANCE_CAUSES = {
    "Machine Problem*",
    "Controller Problem*",
    "Robot Problem*",
    "Scheduled Maintenance*",
    "Power Breakdown (Unscheduled)*",
}

SECTION_CONFIG = {
    "RIP> DPL> Plastic-3": {
        "total_mcs": 61,
        "mapping": {
            "IMM-160-6": "A1-160", "IMM-120-20": "A2-120", "IMM-120-28": "A3-120", "IMM-120-29": "A4-120",
            "IMM-160-7": "A5-160", "IMM-160-12": "A6-160", "IMM-160-48": "A7-160", "IMM-120-11": "B1-120",
            "IMM-120-15": "B2-120", "IMM-120-14": "B3-120", "IMM-120-75": "B4-120", "IMM-90-8": "B5-90PC",
            "IMM-90-9": "B6-90PC", "IMM-120-32": "B7-120PC", "IMM-120-27": "B8-120PC", "IMM-120-4": "C1-120",
            "IMM-160-17": "C2-160", "IMM-120-22": "C3-120", "IMM-120-46": "C4-120PC", "IMM-90-4": "C5-90",
            "IMM-120-47": "C6-120", "IMM-160-51": "C7-160", "IMM-160-39": "D1-160", "IMM-160-79": "D2-160",
            "IMM-160-80": "D3-160", "IMM-280R-25": "A1-280TC", "IMM-380-5": "A2-380", "IMM-380-81": "A3-380 (PC)",
            "IMM-380-80": "A4-380", "IMM-330-4": "A5-HP-330", "IMM-470-5": "B1-470", "IMM-380-6": "B2-380",
            "IMM-530-15": "B3-530", "IMM-530-16": "B4-530", "IMM-530-22": "B5-530", "IMM-380-4": "B6-380",
            "IMM-800-30": "C1-800-30", "IMM-800-31": "C2-800-31", "IMM-270-1": "C3-270-1", "IMM-380-73": "C4-380-73",
            "IMM-380-44": "C5-380-44", "IMM-280R-3": "C6-280TC", "IMM-280R-24": "D1-280TC", "IMM-250-106": "D2-MA2-250",
            "IMM-330-1": "D3-330-1", "IMM-330-5": "D4-HP-330-5", "IMM-428-1": "D5-428-1", "IMM-428-4": "D6-HP-428-4",
            "IMM-330-8": "D7-HP-330", "IMM-380-90": "E1-380-90", "IMM-380-94": "E2-380-94", "IMM-380-88": "E3-380-88",
            "IMM-380-76": "E4-380-76", "IMM-380-62": "E5-380-62", "IMM-380-75": "E6-380-75", "IMM-380-92": "F1-380-92",
            "IMM-380-93": "F2-380-93", "IMM-380-98": "F3-380-98", "IMM-380-99": "F4-380-99", "IMM-380-101": "F5-380-101",
            "IMM-380-100": "F6-380-100",
        },
    },
    "RIP> DPL> Plastic-6.1": {
        "total_mcs": 76,
        "mapping": {
            "IMM-380-86": "A1", "IMM-470-4": "A2", "IMM-380-43": "A3", "IMM-380-13": "A4",
            "IMM-380-29": "A5", "IMM-380-42": "A6", "IMM-380-97": "A7", "IMM-380-3": "B1",
            "IMM-380-57": "B2", "IMM-380-14": "B3", "IMM-380-15": "B4", "IMM-380-46": "B5",
            "IMM-380-7": "B6", "IMM-380-53": "B7", "IMM-380-66": "C1", "IMM-470-16": "C2",
            "IMM-380-70": "C3", "IMM-380-79": "C4", "IMM-380-26": "C5", "IMM-160-61": "C6",
            "IMM-380-30": "C7", "IMM-470-6": "D1", "IMM-470-7": "D2", "IMM-380-74": "D3",
            "IMM-380-12": "D4", "IMM-380-71": "D5", "IMM-380-10": "D6", "IMM-380-72": "D7",
            "IMM-470-17": "E1", "IMM-380-78": "E2", "IMM-380-82": "E3", "IMM-530-7": "E4",
            "IMM-530-6": "E5", "IMM-530-11": "E6", "IMM-500-2": "F1", "IMM-530-19": "F2",
            "IMM-250-29": "F3", "IMM-250-151": "F4", "IMM-250-286": "F5", "IMM-250-18": "F6",
            "IMM-260-2": "F7", "IMM-250-208": "G1", "IMM-250-157": "G2", "IMM-250-284": "G3",
            "IMM-160-78": "G4", "IMM-250-284": "G5", "IMM-160-81": "G6", "IMM-160-72": "G7",
            "IMM-160-77": "G8", "IMM-160-35": "G9", "IMM-160-34": "G10", "IMM-160-42": "G11",
            "IMM-160-68": "H1", "IMM-160-69": "H2", "IMM-160-70": "H3", "IMM-160-71": "H4",
            "IMM-160-73": "H5", "IMM-120-85": "H6", "IMM-90-1": "H7", "IMM-90-2": "H8",
            "IMM-250-175": "I1", "IMM-250-180": "I2", "IMM-250-179": "I3", "IMM-250-14": "I4",
            "IMM-250-15": "I5", "IMM-250-16": "I6", "IMM-250-6": "I7", "IMM-250-8": "I8",
            "IMM-250-97": "J1", "IMM-250-74": "J2", "IMM-250-79": "J3", "IMM-250-47": "J4",
            "IMM-250-60": "J5", "IMM-250-131": "J6", "IMM-250-138": "J7", "IMM-250-168": "J8",
        },
    },
    "RIP> DPL> Plastic-6.2": {
        "total_mcs": 59,
        "mapping": {
            "IMM-250-23": "A1", "IMM-250-100": "A2", "IMM-250-184": "A3", "IMM-160-101": "A4",
            "IMM-160-19": "A5", "IMM-250-35": "A6", "IMM-250-27": "A7", "IMM-250-193": "A8",
            "IMM-250-26": "A9", "IMM-250-98": "B1", "IMM-250-40": "B2", "IMM-250-22": "B3",
            "IMM-120-18": "B4", "IMM-250-173": "B5", "IMM-250-113": "B6", "IMM-160-14": "B7",
            "IMM-160-10": "B8", "IMM-250-19": "B9", "IMM-250-24": "C1", "IMM-250-32": "C2",
            "IMM-250-37": "C3", "IMM-120-26": "C4", "IMM-250-191": "C5", "IMM-120-21": "C6",
            "IMM-250-75": "C7", "IMM-250-88": "C8", "IMM-250-192": "C9", "IMM-250-99": "D1",
            "IMM-250-129": "D2", "IMM-250-190": "D3", "IMM-160-63": "D4", "IMM-250-154": "D5",
            "IMM-250-119": "D6", "IMM-250-189": "D7", "IMM-250-102": "D9", "IMM-250-130": "D10",
            "IMM-250-161": "E1", "IMM-160-64": "E4", "IMM-250-46": "E6", "IMM-250-185": "E7",
            "IMM-250-170": "E8", "IMM-160-98": "F1", "IMM-160-102": "F2", "IMM-160-99": "F3",
            "IMM-160-62": "F4", "IMM-90-15": "F5", "IMM-120-54": "F6", "IMM-90-3": "F7",
            "IMM-90-6": "F8", "IMM-160-103": "F9", "IMM-120-56": "G1", "IMM-120-50": "G2",
            "IMM-120-51": "G3", "IMM-120-52": "G4", "IMM-120-53": "G5", "IMM-90-14": "G6",
            "IMM-120-55": "G7", "IMM-120-24": "G9", "IMM-120-57": "G10",
        },
    },
    "RIP> DPL> Plastic-7.1": {
        "total_mcs": 36,
        "mapping": {
            "IMM-530-17": "A1-530", "IMM-530-12": "A2-530", "IMM-330-6": "B1-330", "IMM-265-1": "B2-265",
            "IMM-268-1": "B3-268", "IMM-380-63": "B4-380", "IMM-270-4": "B5-270", "IMM-530-14": "B6-530",
            "IMM-330-2": "C1-330", "IMM-470-13": "C2-470", "IMM-470-20": "C3-470", "IMM-470-21": "C4-470",
            "IMM-428-2": "C5-428", "IMM-530-20": "D1-530", "IMM-380-40": "D2-380", "IMM-470-15": "D3-470",
            "IMM-470-22": "D4-470", "IMM-428-3": "D5-428", "IMM-470-18": "E1-470", "IMM-470-14": "E2-470",
            "IMM-370-1": "E3-370", "IMM-470-11": "E4-470", "IMM-330-7": "E5-330", "IMM-470-12": "F1-470",
            "IMM-428-5": "F2-428", "IMM-470-19": "G1-470", "IMM-380-77": "G2-380", "IMM-380-67": "H1-380",
            "IMM-330-3": "H2-330", "IMM-380-55": "I1-380", "IMM-380-61": "I2-380", "IMM-270-5": "J1-270",
            "IMM-428-6": "J2-428", "IMM-270-3": "J3-270", "IMM-428-8": "K1-428", "IMM-428-7": "K2-428",
        },
    },
    "RIP> DPL> Plastic-7.2": {
        "total_mcs": 76,
        "mapping": {
            "BMM-05L-11": "BMM-A1", "BMM-05L-04": "BMM-A2", "BMM-05L-03": "BMM-A3", "BMM-05L-06": "BMM-A4",
            "BMM-05L-05": "BMM-A5", "BMM-05L-16": "BMM-A6", "IMM-250-69": "A1", "IMM-250-176": "A2",
            "IMM-250-70": "A3", "IMM-250-177": "A4", "IMM-250-71": "A5", "IMM-250-178": "A6",
            "IMM-250-72": "A7", "IMM-250-179": "A8", "IMM-250-73": "A9", "IMM-250-180": "A10",
            "IMM-250-74": "B1", "IMM-250-75": "B2", "IMM-250-76": "B3", "IMM-250-77": "B4",
            "IMM-250-78": "B5", "IMM-250-79": "B6", "IMM-250-80": "B7", "IMM-250-81": "B8",
            "IMM-250-82": "B9", "IMM-250-83": "B10", "IMM-250-84": "C1", "IMM-250-85": "C2",
            "IMM-250-86": "C3", "IMM-250-87": "C4", "IMM-250-88": "C5", "IMM-250-89": "C6",
            "IMM-250-90": "C7", "IMM-250-91": "C8", "IMM-250-92": "C9", "IMM-250-93": "C10",
            "IMM-250-94": "D1", "IMM-250-95": "D2", "IMM-250-96": "D3", "IMM-250-97": "D4",
            "IMM-250-98": "D5", "IMM-250-99": "D6", "IMM-250-101": "D7", "IMM-250-102": "D8",
            "IMM-250-103": "D9", "IMM-250-104": "D10", "IMM-250-105": "D11", "IMM-160-68": "E1",
            "IMM-160-69": "E2", "IMM-160-70": "E3", "IMM-160-71": "E4", "IMM-160-72": "E5",
            "IMM-160-73": "E6", "IMM-160-74": "E7", "IMM-160-75": "E8", "IMM-160-76": "E9",
            "IMM-160-77": "E10", "IMM-160-78": "E11", "IMM-160-81": "F1", "IMM-160-82": "F2",
            "IMM-160-83": "F3", "IMM-160-84": "F4", "IMM-160-85": "F5", "IMM-160-86": "F6",
            "IMM-120-33": "G1", "IMM-120-34": "G2", "IMM-120-35": "G3", "IMM-120-36": "G4",
            "IMM-90-10": "G5", "IMM-90-11": "G6", "IMM-90-12": "G7", "IMM-260-55": "G8",
            "IMM-160-25": "H1", "IMM-160-41": "H2", "IMM-160-55": "H3", "IMM-160-56": "H4",
        },
    },
    "RIP> DPL> Plastic-7.3": {
        "total_mcs": 77,
        "mapping": {
            "IMM-260-49": "A1-260", "IMM-260-50": "A2-260", "IMM-160-31": "A3-160", "IMM-160-67": "A4-160",
            "IMM-260-51": "A5-260", "IMM-260-52": "A6-260", "IMM-260-53": "A7-260", "IMM-160-60": "A8-160",
            "IMM-160-32": "F1-160", "IMM-160-30": "A10-160", "IMM-260-3": "B1-260", "IMM-260-4": "B2-260",
            "IMM-260-5": "B3-260", "IMM-260-6": "B4-260", "IMM-260-7": "B5-260", "IMM-260-8": "B6-260",
            "IMM-260-9": "B7-260", "IMM-260-10": "B8-260", "IMM-260-11": "B9-260", "IMM-260-12": "B10-260",
            "IMM-260-13": "C1-260", "IMM-260-14": "C2-260", "IMM-250-108": "C2-250", "IMM-250-112": "C3-250",
            "IMM-250-104": "C4-250", "IMM-250-144": "C5-250", "IMM-250-145": "C6-250", "IMM-250-149": "C7-250",
            "IMM-250-150": "C8-250", "IMM-250-140": "C9-250", "IMM-250-139": "C10-250", "IMM-250-132": "C11-250",
            "IMM-250-110": "D1-250", "IMM-250-109": "D2-250", "IMM-250-158": "D3-250", "IMM-250-134": "D4-250",
            "IMM-250-186": "D5-250", "IMM-250-187": "D6-250", "IMM-250-155": "D7-250", "IMM-250-142": "D8-250",
            "IMM-250-285": "D9-250", "IMM-250-283": "D10-250", "IMM-160-54": "D11-160", "IMM-250-288": "E1-250",
            "IMM-250-289": "E2-250", "IMM-250-290": "E3-250", "IMM-250-291": "E4-250", "IMM-250-49": "E5-250",
            "IMM-250-115": "E6-250", "IMM-260-54": "E7-260", "IMM-250-279": "E8-250", "IMM-250-136": "E9-250",
            "IMM-250-135": "E10-250", "IMM-160-22": "E11-160", "IMM-160-27": "F2-160", "IMM-160-28": "F3-160",
            "IMM-160-43": "F4-160", "IMM-160-45": "F5-160", "IMM-160-46": "F6-160", "IMM-160-57": "F7-160",
            "IMM-160-59": "F8-160", "IMM-160-66": "F9-160", "IMM-120-44": "G1-120", "IMM-120-45": "G2-120",
            "IMM-260-55": "G3-260", "IMM-260-21": "G4-260", "IMM-260-22": "G5-260", "IMM-260-23": "G6-260",
            "IMM-260-24": "G7-260", "IMM-260-25": "G8-260", "IMM-260-15": "G9-260", "IMM-260-16": "G10-260",
            "IMM-260-17": "G7-260", "IMM-260-18": "G8-260", "IMM-260-19": "G9-260", "IMM-260-20": "G10-260",
            "IMM-160-44": "G11-160",
        },
    },
}

DEFAULT_SECTION = "RIP> DPL> Plastic-3"


def parse_machine_size(smart_manu_or_pos):
    text = str(smart_manu_or_pos).strip().upper()
    parts = text.split("-")
    if len(parts) >= 2:
        raw_size = parts[1]
        clean_size = raw_size.replace("R", "").replace("TC", "").strip()
        return clean_size if clean_size else raw_size
    return "Other"


def resolve_section_config(df=None, fallback_name=DEFAULT_SECTION):
    section_name = fallback_name
    if df is not None and "Section" in df.columns:
        valid_sections = [s for s in df["Section"].dropna().unique() if str(s).strip()]
        for vs in valid_sections:
            for registered in SECTION_CONFIG.keys():
                if registered.lower() in str(vs).lower() or str(vs).lower() in registered.lower():
                    section_name = registered
                    break

    cfg = SECTION_CONFIG.get(section_name, SECTION_CONFIG[DEFAULT_SECTION])
    total_mcs = cfg["total_mcs"]
    daily_avail_hrs = total_mcs * 24.0
    pos_map = cfg["mapping"]

    sizes_series = [parse_machine_size(sm) for sm in pos_map.keys()]
    size_counts = {}
    for s in sizes_series:
        size_counts[s] = size_counts.get(s, 0) + 1

    unique_sizes = sorted(size_counts.keys(), key=lambda x: (len(x), x), reverse=True)
    return section_name, total_mcs, daily_avail_hrs, pos_map, size_counts, unique_sizes


TOTAL_PLANT_MCS = SECTION_CONFIG[DEFAULT_SECTION]["total_mcs"]
DAILY_AVAILABLE_HRS = TOTAL_PLANT_MCS * 24.0
POS_MAP = SECTION_CONFIG[DEFAULT_SECTION]["mapping"]
LINE_MAP = {k: "-" for k in POS_MAP.keys()}
EXCEL_SIZES = ["160", "90", "120", "250", "270", "280", "380", "330", "470", "530", "800", "428"]


# =========================================================
# CENTRAL NPT CAUSE CATEGORIZATION REGISTRY (65 CAUSES)
# Grouped by Industrial Maintenance & Operational Ownership
# =========================================================
NPT_CATEGORIES = {
    "Planning & Commercial Idle": [
        "No Demand",
        "No Demand*",
        "Over Stock*",
        "Manpower Short*",
        "Holiday*",
        "Shutdown (Scheduled)",
        "One Machine Idle for Another Machine (Announced)",
        "Audit",
        "Line Balancing*",
        "Machine Transfer*",
        "Inventory (Raw Materials/Finished Goods/In Process)*",
    ],
    "Machine & Technical Breakdown": [
        "Machine Problem*",
        "Robot Problem*",
        "Controller Problem*",
        "RMCS Problem*",
        "Oil or water Leakage*",
        "Barrel Heater Problem*",
        "Heater problem*",
        "Heater  problem*",
        "Machine Heater Problem*",
        "Lubricant fail*",
        "Scheduled Maintenance*",
        "Machine Greasing*",
        "Blowing machine Gripper Problem*",
        "Blow pin problem*",
        "Head lift problem*",
        "Head polish*",
        "Twin Barrel Problem*",
        "Cutter Problem*",
        "Cutting problem",
        "Unplanned Downtime",
        "Alternative Problem*",
    ],
    "Mold & Tooling Issues": [
        "Mold Problem*",
        "Mold Insert Change*",
        "Head Change*",
        "Barrel Change*",
    ],
    "SMED & Changeover": [
        "Mold Change*",
        "Color Change*",
        "Item Change*",
        "Parison Adjusting*",
        "Pre-Heating Time",
        "CIP/SIP/COP",
    ],
    "Operational & Quality Loss": [
        "Product Jam*",
        "Nozzle Jam*",
        "Hopper Jam*",
        "Color variation*",
        "Defective product*",
        "Over Flash",
        "Moisture Problem*",
        "Material Leakage*",
        "Material Shortage*",
        "Accessories Shortage*",
        "Curve problem*",
        "View strip problem*",
        "Sample + Mold Test (RND)*",
        "Machine Cleaning Break*",
        "Prayer Break*",
        "Lunch/ Dinner Break",
        "Meal Break*",
        "Sehri/ Iftar Time",
        "Fire Drill",
    ],
    "Utilities & Facilities": [
        "Power Breakdown (Unscheduled)*",
        "Power Breakdown (Scheduled)*",
        "Utility (Air, Water & Crane)*",
        "Utility problem*",
        "Server Error (Announced)",
        "Server Error (Unannounced)",
    ],
}

# Fast O(1) lookup dictionary for row-level assignment (normalizes asterisks & casing)
NPT_CAUSE_TO_CATEGORY = {}
for category_name, causes_list in NPT_CATEGORIES.items():
    for c_item in causes_list:
        clean_key = str(c_item).replace("*", "").strip().lower()
        NPT_CAUSE_TO_CATEGORY[clean_key] = category_name


def resolve_npt_category(cause_str):
    """
    Resolves any raw NPT cause string to its industrial department/category.
    Matches variations with/without asterisks, whitespace, or case differences.
    """
    if not cause_str or str(cause_str).strip() in ["", "nan", "None"]:
        return "Unassigned / Pending Log"
    clean_k = str(cause_str).replace("*", "").strip().lower()
    return NPT_CAUSE_TO_CATEGORY.get(clean_k, "Other Operational Loss")
