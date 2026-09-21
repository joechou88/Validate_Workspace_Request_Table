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

EXPECTED_DATATYPES = {
    "annual_financials_global": "WC06105,WC06001,WC06004,IBTKR,WC05601,WC06006,WC06008,WC06026,WC06010,WC07021,WC07022,WC07023,WC07024,WC07025,WC07026,WC07027,WC07028,WC05350,WC05427,WC11496, WC07536,X(WC01801)~U$,WC07546,WC07800,WC07805,WC07810,WC05200,X(WC18545)~U$,X(WC18546)~U$,X(WC18547)~U$,X(WC18548)~U$,X(WC02999)~U$,X(WC08001)~U$,X(WC08002)~U$,X(WC03351)~U$,X(WC03255)~U$,X(WC05476)~U$,X(WC05001)~U$,X(WC03501)~U$,X(WC03995)~U$,X(WC01551)~U$,X(WC18198)~U$,X(WC01401)~U$,X(WC01751)~U$,X(WC02201)~U$,X(WC03101)~U$,X(WC02051)~U$,X(WC02101)~U$,X(WC01253)~U$,X(WC01254),X(WC04860)~U$,X(WC01001)~U$,X(WC02051)~U$,X(WC02301)~U$,X(WC02501)~U$,WC19506,WC19516,WC19526,WC19536,WC19546,WC19556,WC19566,WC19576,WC19586,WC19596,WC19600,WC19610,WC19620,WC19630,WC19640,WC19650,WC19660,WC19670,WC19680,WC19690,X(WC01451)~U$,X(WC18186)~U$,X(WC18187)~U$,X(WC18188)~U$,X(WC18189)~U$,X(WC03263)~U$,X(WC18166)~U$,X(WC04199)~U$,X(WC03063)~U$,X(WC04551)~U$,X(WC05001)~U$,WC05301,X(WC01075)~U$,X(WC01251)~U$,X(WC07151)~U$,X(WC07101)~U$,X(WC07126)~U$,WC11559,X(WC04500)~U$,X(WC04890)~U$,X(WC04401)~U$,X(WC04302)~U$,X(WC11496)~U$,WC07011,WC07815,WC11556,WC11557,WC11558,X(WC18705)~U$,X(WC18706)~U$,X(WC18707)~U$,X(WC18708)~U$,X(WC18709)~U$,X(WC18710)~U$,X(WC18713)~U$,X(WC18714)~U$,X(WC18715)~U$,X(WC02518)~U$,X(WC02519)~U$,X(WC02520)~U$,X(WC18721)~U$,X(WC18725)~U$,X(WC18722)~U$,X(WC18147)~U$,X(WC01158)~U$,X(WC01159)~U$,X(WC18099)~U$, X(WC04601)~US,X(WC01201)~US",
    "annual_financials_insurance": "WC06105,WC06001,WC06004,IBTKR,WC05601,WC06006,WC06008,WC06026,WC06010,WC07021,WC07022,WC07023,WC07024,WC07025,WC07026,WC07027,WC07028,WC05350,WC05427,WC07800,WC07805,WC07810,WC11496,WC11556,WC11557,WC11558,WC11559,X(WC18871)~U$,X(WC18872)~U$,X(WC18873)~U$,X(WC18874)~U$,X(WC18875)~U$,X(WC18876)~U$,X(WC18881)~U$,X(WC18882)~U$,X(WC18883)~U$,X(WC01002)~U$,X(WC01004)~U$,X(WC01071)~U$,X(WC03005)~U$,X(WC18204)~U$,X(WC18203)~U$,X(WC02245)~U$,X(WC02999)~U$,X(WC03351)~U$,X(WC03501)~U$,X(WC01551)~U$,X(WC04860)~U$,X(WC05201)~U$,WC08326,WC08301,WC05301,WC07536,WC07041",
    "daily_return_global": {
        7: "WC06105, WC06001, WC06004, WC05601, WC06038, WC06006, IBTKR,WC06008, WC06026, WC05350, WC07021",
        8: "X(RI)~U$",
        9: "X(LI)~U$",
        10: "X(MV)~U$",
        11: "WC05301",
        12: "VO"
    },
    "daily_return_industry": {
        7: "X(RI)~U$"
    },
    "quarterly_financials_insurance": {
        7: "WC06105,WC06001,WC06004,IBTKR,WC05601,WC06006,WC06008,WC06026,WC06010,WC07021,WC07022,WC07023,WC07024,WC07025,WC07026,WC07027,WC07028,WC05350,WC05427,WC07800,WC07805,WC07810,WC11496,WC11556,WC11557,WC11558,WC11559",
        8: "X(WC18871)~U$,X(WC18872)~U$,X(WC18873)~U$,X(WC18874)~U$,X(WC18875)~U$,X(WC18876)~U$,X(WC18881)~U$,X(WC18882)~U$,X(WC18883)~U$,X(WC01002)~U$,X(WC01004)~U$,X(WC01071)~U$,X(WC03005)~U$,X(WC18204)~U$,X(WC18203)~U$,X(WC02245)~U$,X(WC02999)~U$,X(WC03351)~U$,X(WC03501)~U$,X(WC01551)~U$,X(WC04860)~U$,X(WC05201)~U$,WC08326,WC08301,WC05301,WC07536,WC07041"
    }
}
