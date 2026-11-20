
import csv
import difflib
import re
import sys
import unicodedata
from pathlib import Path
from datetime import date, timedelta
from decimal import Decimal

# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------

PROJECT_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = (
    Path(sys.argv[1]) if len(sys.argv) > 1
    else PROJECT_DIR / "data" / "raw" / "leads_test_messy.csv"
)

OUTPUT_FILE = (
    Path(sys.argv[2]) if len(sys.argv) > 2
    else PROJECT_DIR / "data" / "processed" / "leads_clean.csv"
)

IN_COLS = ["lead_id", "full_name", "email", "phone", "instagram_handle", "city", "state",
           "country", "institution", "role", "source", "date_added", "followers"]
OUT_COLS = ["lead_id", "first_name", "last_name", "email", "phone", "instagram_handle", "city",
            "state", "country", "institution", "role", "source", "date_added", "followers", "flags"]

LAST_ALLOWED_DATE = date(2026, 11, 11)   # Rule 13: later dates are invalid
EXCEL_EPOCH = date(1899, 12, 30)         # Rule 13: Excel serial 1 = 1899-12-31
MAX_FOLLOWERS = 5_000_000                # Rule 14

PLACEHOLDERS = {"", "n/a", "na", "none", "null", "-"}   # "no value" markers

# ----------------------------------------------------------------------------
# Canonical lists (Rules 7, 8, 10, 11, 12)
# ----------------------------------------------------------------------------
CITIES = ["Mumbai", "Navi Mumbai", "Thane", "Pune", "Delhi", "Bengaluru", "Hyderabad",
          "Chennai", "Kolkata", "Ahmedabad", "Jaipur"]
CITY_SYNONYMS = {"Bombay": "Mumbai", "Poona": "Pune", "New Delhi": "Delhi",
                 "Bangalore": "Bengaluru", "BLR": "Bengaluru", "Hyd": "Hyderabad",
                 "Madras": "Chennai", "Calcutta": "Kolkata"}

STATES = ["Maharashtra", "Delhi", "Karnataka", "Telangana", "Tamil Nadu", "West Bengal",
          "Gujarat", "Rajasthan"]
STATE_SYNONYMS = {"MH": "Maharashtra", "DL": "Delhi", "KA": "Karnataka", "TG": "Telangana",
                  "TN": "Tamil Nadu", "WB": "West Bengal", "GJ": "Gujarat", "RJ": "Rajasthan",
                  "NCT Delhi": "Delhi"}

# Rule 8: if the state is empty, take it from the cleaned city
CITY_TO_STATE = {"Mumbai": "Maharashtra", "Navi Mumbai": "Maharashtra", "Thane": "Maharashtra",
                 "Pune": "Maharashtra", "Delhi": "Delhi", "Bengaluru": "Karnataka",
                 "Hyderabad": "Telangana", "Chennai": "Tamil Nadu", "Kolkata": "West Bengal",
                 "Ahmedabad": "Gujarat", "Jaipur": "Rajasthan"}

INSTITUTIONS = ["Lotus Film Academy", "Lotus Film Academy, Andheri", "Marine Drive School of Cinema",
                "Sahyadri Institute of Media Studies", "Gateway College of Arts",
                "Aarambh Theatre Institute", "Western Line Film School", "Bandra Creative Institute",
                "Silver Reel Academy", "Juhu Media College", "Kalaa Mandir Institute",
                "Horizon Film and TV Institute", "Orchid School of Visual Arts"]

ROLES = ["Student", "Alumnus", "Faculty", "Freelancer", "Professional"]
ROLE_SYNONYMS = {"Stud.": "Student", "Alumni": "Alumnus", "Ex-student": "Alumnus",
                 "Teacher": "Faculty", "Professor": "Faculty", "Lecturer": "Faculty",
                 "Freelance": "Freelancer", "Free-lancer": "Freelancer",
                 "Working Professional": "Professional"}

SOURCES = ["Institution list", "Event", "Referral", "Social search"]
SOURCE_SYNONYMS = {"Inst. List": "Institution list", "Institution-list": "Institution list",
                   "Events": "Event", "Ref": "Referral", "Refer": "Referral",
                   "Social": "Social search", "Social-search": "Social search"}

EMAIL_DOMAIN_FIXES = {"gmial.com": "gmail.com", "gamil.com": "gmail.com", "gmail.con": "gmail.com",
                      "yahoo.con": "yahoo.com", "hotmial.com": "hotmail.com"}

HONORIFICS = {"mr", "mrs", "ms", "dr"}

MONTH_NAMES = ["january", "february", "march", "april", "may", "june", "july", "august",
               "september", "october", "november", "december"]

# Values we could not map to a canonical value are collected here and printed at the end
unmatched = []


# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------
def collapse(text):
    """Trim and replace any run of whitespace by one space."""
    return re.sub(r"\s+", " ", text).strip()


def norm(text):
    """Comparison key: lowercase, no '-', '_' or '.', single spaces."""
    return collapse(re.sub(r"[-_.]", " ", text)).lower()


def build_choices(canonical, synonyms):
    """dict  normalised text -> canonical value  (canonical names and synonyms)."""
    choices = {}
    for name in canonical:
        choices[norm(name)] = name
    for word, name in synonyms.items():
        choices[norm(word)] = name
    return choices


CITY_CHOICES = build_choices(CITIES, CITY_SYNONYMS)
STATE_CHOICES = build_choices(STATES, STATE_SYNONYMS)
INSTITUTION_CHOICES = build_choices(INSTITUTIONS, {})
ROLE_CHOICES = build_choices(ROLES, ROLE_SYNONYMS)
SOURCE_CHOICES = build_choices(SOURCES, SOURCE_SYNONYMS)


def match_canonical(raw, choices, cutoff=0.75):
    """
    Map a raw text to a canonical value.
    1) exact match after normalising (fixes case, spaces, synonyms, codes)
    2) otherwise the nearest spelling (difflib), if it is similar enough
    Returns "" when there is no value or no good match.
    """
    key = norm(raw)
    if key in PLACEHOLDERS:
        return ""
    if key in choices:
        return choices[key]
    if len(key) >= 4:                       # never fuzzy-match very short strings
        close = difflib.get_close_matches(key, list(choices), n=1, cutoff=cutoff)
        if close:
            return choices[close[0]]
    unmatched.append(raw)
    return ""


# ----------------------------------------------------------------------------
# Rule 3 - name
# ----------------------------------------------------------------------------
def keep_char(ch):
    """Keep letters, accents, space, apostrophe, hyphen, dot and comma. Drops emoji and symbols."""
    category = unicodedata.category(ch)
    if category.startswith("L"):
        return True
    if category.startswith("M") and not (0xFE00 <= ord(ch) <= 0xFE0F):  # accents yes, emoji selector no
        return True
    return ch in " '-.,"


def title_case(word):
    """isha -> Isha, anne-marie -> Anne-Marie."""
    return "-".join(part.capitalize() for part in word.split("-"))


def clean_name(raw):
    text = collapse("".join(ch for ch in raw if keep_char(ch)))     # remove emoji, tidy spaces
    if "," in text:                                                 # "Last, First" -> "First Last"
        last, first = text.split(",", 1)
        text = first + " " + last
    words = collapse(text.replace(",", " ")).split(" ")
    words = [w for w in words if w and w.strip(".").lower() not in HONORIFICS]   # drop Mr. Mrs. Ms. Dr.
    words = [title_case(w) for w in words]
    first_name = words[0] if words else ""
    last_name = " ".join(words[1:])
    return first_name, last_name


# ----------------------------------------------------------------------------
# Rule 4 - email
# ----------------------------------------------------------------------------
EMAIL_RE = re.compile(r"^[a-z0-9._%+\-]+@([a-z0-9\-]+\.)+[a-z]{2,}$")


def clean_email(raw):
    """Returns (email, is_invalid)."""
    text = raw.split(";")[0].strip().lower()        # several emails -> keep the first
    if text == "":
        return "", False                            # empty stays empty, no flag
    if "@" in text:
        local, _, domain = text.partition("@")
        text = local + "@" + EMAIL_DOMAIN_FIXES.get(domain, domain)    # fix domain typos
    if EMAIL_RE.match(text):
        return text, False
    return "", True                                 # no @, no dot in domain, N/A, none, -, ...


# ----------------------------------------------------------------------------
# Rule 5 - phone
# ----------------------------------------------------------------------------
def clean_phone(raw):
    """Returns (+91XXXXXXXXXX, is_invalid)."""
    text = raw.strip()
    if text == "":
        return "", False
    text = re.sub(r"[\s\-()]", "", text)            # remove spaces, dashes, brackets
    if text.startswith("+"):
        text = text[1:]
    if not re.fullmatch(r"[0-9]+", text):           # letters, '.', 'E+09' (scientific notation), NA ...
        return "", True
    if len(text) == 12 and text.startswith("91"):   # leading +91 / 91
        text = text[2:]
    elif len(text) == 11 and text.startswith("0"):  # leading 0
        text = text[1:]
    if len(text) == 10 and text[0] in "6789" and text not in ("0000000000", "1234567890"):
        return "+91" + text, False
    return "", True


# ----------------------------------------------------------------------------
# Rule 6 - instagram handle
# ----------------------------------------------------------------------------
def clean_instagram(raw):
    text = raw.strip().lower()
    text = re.split(r"[?#]", text)[0]               # drop query string
    if "instagram.com/" in text:                    # drop URL part (http(s)://, www.)
        text = text.split("instagram.com/", 1)[1]
    text = text.strip("/ ").split("/")[0]           # first path segment only
    return text.lstrip("@").strip()


# ----------------------------------------------------------------------------
# Rule 13 - date
# ----------------------------------------------------------------------------
def month_number(word):
    """'Nov', 'nov.', 'November', 'Sept' -> 11 (None if not a month)."""
    word = word.lower().strip(".")
    if len(word) < 3:
        return None
    for number, name in enumerate(MONTH_NAMES, start=1):
        if name.startswith(word):
            return number
    return None


def parse_date(raw):
    """Returns (YYYY-MM-DD, is_invalid)."""
    text = collapse(raw)
    if text == "":
        return "", False
    result = None
    try:
        if re.fullmatch(r"\d{5}", text):                                    # Excel serial number
            result = EXCEL_EPOCH + timedelta(days=int(text))
        elif re.fullmatch(r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2}", text):          # YYYY/MM/DD, YYYY-MM-DD
            y, m, d = re.split(r"[-/.]", text)
            result = date(int(y), int(m), int(d))
        elif re.fullmatch(r"\d{1,2}[-/.]\d{1,2}[-/.]\d{4}", text):          # DD/MM/YYYY (day first)
            d, m, y = re.split(r"[-/.]", text)
            result = date(int(y), int(m), int(d))
        else:
            m1 = re.fullmatch(r"(\d{1,2}) ([A-Za-z]+\.?),? (\d{4})", text)      # 3 Nov 2026
            m2 = re.fullmatch(r"([A-Za-z]+\.?) (\d{1,2}),? (\d{4})", text)       # Nov 3, 2026
            if m1 and month_number(m1.group(2)):
                result = date(int(m1.group(3)), month_number(m1.group(2)), int(m1.group(1)))
            elif m2 and month_number(m2.group(1)):
                result = date(int(m2.group(3)), month_number(m2.group(1)), int(m2.group(2)))
    except (ValueError, OverflowError):             # 31/02/2026, 45/09/2026, ...
        result = None
    if result is None or result > LAST_ALLOWED_DATE:
        return "", True
    return result.isoformat(), False


# ----------------------------------------------------------------------------
# Rule 14 - followers
# ----------------------------------------------------------------------------
def clean_followers(raw):
    """Returns (integer as text, is_invalid)."""
    text = collapse(raw).lower()
    if text == "":
        return "", False
    text = re.sub(r" ?followers?$", "", text)       # "123 followers" -> "123"
    number = None
    if re.fullmatch(r"\d+", text) or re.fullmatch(r"\d{1,3}(,\d{3})+", text):   # 1234 or 1,234
        number = int(text.replace(",", ""))
    elif re.fullmatch(r"\d+(\.\d+)?k", text):                                   # 1.2k
        number = int(round(Decimal(text[:-1]) * 1000))
    if number is None or number > MAX_FOLLOWERS:    # negative, text, too big
        return "", True
    return str(number), False


# ----------------------------------------------------------------------------
# Rule 1 - structure
# ----------------------------------------------------------------------------
def repair_row(fields):
    """
    A row with too many fields has an unquoted comma inside one value.
    Try joining every pair of neighbouring fields and keep the version where the
    role, the source and the institution are recognisable.
    """
    while len(fields) > len(IN_COLS):
        best_score, best_fields = -1, None
        for i in range(len(fields) - 1):
            joined = fields[:i] + [",".join(fields[i:i + 2])] + fields[i + 2:]
            score = 0
            if norm(joined[9]) in ROLE_CHOICES:
                score += 1
            if norm(joined[10]) in SOURCE_CHOICES:
                score += 1
            if norm(joined[8]) in INSTITUTION_CHOICES or norm(joined[8]) in PLACEHOLDERS:
                score += 1
            if score > best_score:                  # '>' keeps the first best candidate
                best_score, best_fields = score, joined
        fields = best_fields
    return fields


def read_rows(path):
    """Read the file and return a list of 13-field rows (empty rows, repeated headers, footer removed)."""
    rows = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        for fields in csv.reader(f):
            if not any(x.strip() for x in fields):                      # empty row
                continue
            first = fields[0].strip().lower()
            if first == "lead_id":                                      # header (first or repeated)
                continue
            if first.startswith("total rows"):                          # footer line
                continue
            if len(fields) > len(IN_COLS):                              # too many fields -> repair
                fields = repair_row(fields)
            fields = fields + [""] * (len(IN_COLS) - len(fields))       # too few -> pad
            rows.append(fields)
    return rows


# ----------------------------------------------------------------------------
# Clean one row
# ----------------------------------------------------------------------------
def clean_row(fields):
    raw = dict(zip(IN_COLS, fields))
    flags = set()

    first_name, last_name = clean_name(raw["full_name"])

    email, bad = clean_email(raw["email"])
    if bad:
        flags.add("invalid_email")

    phone, bad = clean_phone(raw["phone"])
    if bad:
        flags.add("invalid_phone")

    city = match_canonical(raw["city"], CITY_CHOICES)

    state = match_canonical(raw["state"], STATE_CHOICES)
    if state == "":
        state = CITY_TO_STATE.get(city, "")         # Rule 8: derive from the cleaned city

    date_added, bad = parse_date(raw["date_added"])
    if bad:
        flags.add("invalid_date")

    followers, bad = clean_followers(raw["followers"])
    if bad:
        flags.add("invalid_followers")

    return {
        "lead_id": raw["lead_id"].strip(),
        "first_name": first_name,
        "last_name": last_name,
        "email": email,
        "phone": phone,
        "instagram_handle": clean_instagram(raw["instagram_handle"]),
        "city": city,
        "state": state,
        "country": "India",                         # Rule 9: always India
        "institution": match_canonical(raw["institution"], INSTITUTION_CHOICES),
        "role": match_canonical(raw["role"], ROLE_CHOICES),
        "source": match_canonical(raw["source"], SOURCE_CHOICES),
        "date_added": date_added,
        "followers": followers,
        "flags": ";".join(sorted(flags)),           # Rule 15: alphabetical, ';' separated
    }


# ----------------------------------------------------------------------------
# Rule 2 - duplicates
# ----------------------------------------------------------------------------
def remove_duplicates(records):
    """
    Group rows that share a valid email, phone or instagram handle (a chain A-B, B-C
    puts A, B and C in one group). From each group keep ONE row: the earliest date_added;
    if the dates are equal or one is unknown, the one that comes first in the file.
    Nothing is merged from the removed rows.
    """
    parent = list(range(len(records)))              # union-find: parent[i] = representative of i's group

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    first_seen = {}                                 # (kind, value) -> index of the first row having it
    for i, rec in enumerate(records):
        for kind in ("email", "phone", "instagram_handle"):
            value = rec[kind]
            if value == "":                         # only compare non-empty (= valid) values
                continue
            key = (kind, value)
            if key in first_seen:
                parent[find(i)] = find(first_seen[key])
            else:
                first_seen[key] = i

    groups = {}
    for i in range(len(records)):                   # i grows, so each group list is in file order
        groups.setdefault(find(i), []).append(i)

    keep = set()
    removed_info = []
    for members in groups.values():
        winner = members[0]                         # default: first in file
        for i in members[1:]:
            a, b = records[i]["date_added"], records[winner]["date_added"]
            if a and b and a < b:                   # strictly earlier known date wins
                winner = i
        keep.add(winner)
        for i in members:
            if i != winner:
                removed_info.append((records[i]["lead_id"], records[winner]["lead_id"]))
    return [rec for i, rec in enumerate(records) if i in keep], removed_info


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def main():

    # Read input CSV
    rows = read_rows(INPUT_FILE)

    records = [
        clean_row(fields)
        for fields in rows
    ]

    # Remove duplicates
    final, removed_info = remove_duplicates(records)

    # Create output directory if it does not exist
    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    # Save final CSV
    with open(
        OUTPUT_FILE,
        "w",
        newline="",
        encoding="utf-8"
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=OUT_COLS
        )

        writer.writeheader()
        writer.writerows(final)

    # Display summary
    print(f"Data rows read        : {len(rows)}")
    print(f"Duplicates removed    : {len(removed_info)}")

    for removed, kept in removed_info:
        print(f"   removed {removed} (kept {kept})")

    print(f"Rows written          : {len(final)} -> {OUTPUT_FILE}")

    print(f"Output columns        : {len(OUT_COLS)}")

    if unmatched:
        print(
            "Values with no canonical match (left empty):",
            sorted(set(unmatched))
        )


if __name__ == "__main__":
    main()

