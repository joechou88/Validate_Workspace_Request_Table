import re

# Verbose regex for filenames with named groups
FILENAME_PATTERN = re.compile(
    r"""
    ^
    (?:
        (?P<country>[A-Z][A-Za-z]+(?:-[A-Z][A-Za-z]+)*)  # Match countries (e.g., Germany, South-Korea)
        |
        (?P<category>[A-Za-z0-9\-]+)                                     # Match categories (e.g., financial-return, non-life-insurance-1)
    )
    (?P<company>\d+)?               # Optional digit(s) for company number
    [_-]                            # Separator (underscore or hyphen)
    (?:(?!\d{4}(?:[-_]|$)).*?[_-])? # Optional intermediate text (e.g., quarterly-financial_), avoiding start year consumption
    (?P<start_year>\d{4})           # 4-digit start year
    (?:[-_](?P<end_year>\d{4}))?    # Optional hyphen/underscore and 4-digit end year
    (?:_(?P<suffix>[A-Za-z]+))?     # Optional suffix, e.g., "_A"
    (?:\.xlsm)?$                    # File extension
    """,
    re.VERBOSE
)

# Simple regex to extract a 4-digit year from cell values (e.g., '2018-01-01' or '10/12/2023')
CELL_YEAR_PATTERN = re.compile(r'((?:19|20)\d{2})')

EXCLUDED_FOLDERS = {'insurance_company_list', 'annual_audit_fee', '__pycache__'}

COUNTRY_MNEMONIC = {
    "Argentina": "FARALL",
    "Australia": "FAUALL",
    "Austria": "FATALL",
    "Belgium": "FBEALL",
    "Brazil": "FBRALL",
    "Canada": "FCAALL",
    "China": "FCNALL",
    "Columbia": "FCOALL",
    "Denmark": "FDKALL",
    "Finland": "FFIALL",
    "France": "FFRALL",
    "Germany1": "FDEALL1",
    "Germany2": "FDEALL2",
    "Germany3": "FDEALL3",
    "Germany4": "FDEALL4",
    "Germany5": "FDEALL5",
    "Germany6": "FDEALL6",
    "Germany7": "FDEALL7",
    "Germany8": "FDEALL8",
    "Greece": "FGRALL",
    "Hong-Kong": "FHKALL",
    "Hungary": "FHUALL",
    "India": "FINALL",
    "Indonesia": "FIDALL",
    "Ireland": "FIEALL",
    "Israel": "FILALL",
    "Italy": "FITALL",
    "Japan": "FJPALL",
    "Malaysia": "FMYALL",
    "Mexico": "FMXALL",
    "Netherlands": "FNLALL",
    "New-Zealand": "FNZALL",
    "Norway": "FNOALL",
    "Pakistan": "FPKALL",
    "Peru": "FPEALL",
    "Philippines": "FPHALL",
    "Poland": "FPLALL",
    "Portugal": "FPTALL",
    "Singapore": "FSGALL",
    "South-Korea": "FKRALL",
    "Spain": "FESALL",
    "Sweden": "FSEALL",
    "Switzerland": "FCHALL",
    "Taiwan": "FTWALL",
    "Thailand": "FTHALL",
    "Turkey": "FTRALL",
    "UK": "FGBALL",
    "Vietnam": "FVNALL"
}
