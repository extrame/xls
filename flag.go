package xls

// Original file header of ParseXL (used as the base for this class):
// --------------------------------------------------------------------------------
// Adapted from Excel_Spreadsheet_Reader developed by users bizon153,
// trex005, and mmp11 (SourceForge.net)
// https://sourceforge.net/projects/phpexcelreader/
// Primary changes made by canyoncasa (dvc) for ParseXL 1.00 ...
//     Modelled moreso after Perl Excel Parse/Write modules
//     Added Parse_Excel_Spreadsheet object
//         Reads a whole worksheet or tab as row,column array or as
//         associated hash of indexed rows and named column fields
//     Added variables for worksheet (tab) indexes and names
//     Added an object call for loading individual woorksheets
//     Changed default indexing defaults to 0 based arrays
//     Fixed date/time and percent formats
//     Includes patches found at SourceForge...
//         unicode patch by nobody
//         unpack("d") machine depedency patch by matchy
//         boundsheet utf16 patch by bjaenichen
//     Renamed functions for shorter names
//     General code cleanup and rigor, including <80 column width
//     Included a testcase Excel file and PHP example calls
//     Code works for PHP 5.x

// Primary changes made by canyoncasa (dvc) for ParseXL 1.10 ...
// http://sourceforge.net/tracker/index.php?func=detail&aid=1466964&group_id=99160&atid=623334
//     Decoding of formula conditions, results, and tokens.
//     Support for user-defined named cells added as an array "namedcells"
//         Patch code for user-defined named cells supports single cells only.
//         NOTE: this patch only works for BIFF8 as BIFF5-7 use a different
//         external sheet reference structure

// ParseXL definitions
const XLS_BIFF8 = 0x0600
const XLS_BIFF7 = 0x0500
const XLS_WorkbookGlobals = 0x0005
const XLS_Worksheet = 0x0010

// record identifiers
const XLS_Type_FORMULA = 0x0006
const XLS_Type_EOF = 0x000a
const XLS_Type_PROTECT = 0x0012
const XLS_Type_OBJECTPROTECT = 0x0063
const XLS_Type_SCENPROTECT = 0x00dd
const XLS_Type_PASSWORD = 0x0013
const XLS_Type_HEADER = 0x0014
const XLS_Type_FOOTER = 0x0015
const XLS_Type_EXTERNSHEET = 0x0017
const XLS_Type_DEFINEDNAME = 0x0018
const XLS_Type_VERTICALPAGEBREAKS = 0x001a
const XLS_Type_HORIZONTALPAGEBREAKS = 0x001b
const XLS_Type_NOTE = 0x001c
const XLS_Type_SELECTION = 0x001d
const XLS_Type_DATEMODE = 0x0022
const XLS_Type_EXTERNNAME = 0x0023
const XLS_Type_LEFTMARGIN = 0x0026
const XLS_Type_RIGHTMARGIN = 0x0027
const XLS_Type_TOPMARGIN = 0x0028
const XLS_Type_BOTTOMMARGIN = 0x0029
const XLS_Type_PRINTGRIDLINES = 0x002b
const XLS_Type_FILEPASS = 0x002f
const XLS_Type_FONT = 0x0031
const XLS_Type_CONTINUE = 0x003c
const XLS_Type_PANE = 0x0041
const XLS_Type_CODEPAGE = 0x0042
const XLS_Type_DEFCOLWIDTH = 0x0055
const XLS_Type_OBJ = 0x005d
const XLS_Type_COLINFO = 0x007d
const XLS_Type_IMDATA = 0x007f
const XLS_Type_SHEETPR = 0x0081
const XLS_Type_HCENTER = 0x0083
const XLS_Type_VCENTER = 0x0084
const XLS_Type_SHEET = 0x0085
const XLS_Type_PALETTE = 0x0092
const XLS_Type_SCL = 0x00a0
const XLS_Type_PAGESETUP = 0x00a1
const XLS_Type_MULRK = 0x00bd
const XLS_Type_MULBLANK = 0x00be
const XLS_Type_DBCELL = 0x00d7
const XLS_Type_XF = 0x00e0
const XLS_Type_MERGEDCELLS = 0x00e5
const XLS_Type_MSODRAWINGGROUP = 0x00eb
const XLS_Type_MSODRAWING = 0x00ec
const XLS_Type_SST = 0x00fc
const XLS_Type_LABELSST = 0x00fd
const XLS_Type_EXTSST = 0x00ff
const XLS_Type_EXTERNALBOOK = 0x01ae
const XLS_Type_DATAVALIDATIONS = 0x01b2
const XLS_Type_TXO = 0x01b6
const XLS_Type_HYPERLINK = 0x01b8
const XLS_Type_DATAVALIDATION = 0x01be
const XLS_Type_DIMENSION = 0x0200
const XLS_Type_BLANK = 0x0201
const XLS_Type_NUMBER = 0x0203
const XLS_Type_LABEL = 0x0204
const XLS_Type_BOOLERR = 0x0205
const XLS_Type_STRING = 0x0207
const XLS_Type_ROW = 0x0208
const XLS_Type_INDEX = 0x020b
const XLS_Type_ARRAY = 0x0221
const XLS_Type_DEFAULTROWHEIGHT = 0x0225
const XLS_Type_WINDOW2 = 0x023e
const XLS_Type_RK = 0x027e
const XLS_Type_STYLE = 0x0293
const XLS_Type_FORMAT = 0x041e
const XLS_Type_SHAREDFMLA = 0x04bc
const XLS_Type_BOF = 0x0809
const XLS_Type_SHEETPROTECTION = 0x0867
const XLS_Type_RANGEPROTECTION = 0x0868
const XLS_Type_SHEETLAYOUT = 0x0862
const XLS_Type_XFEXT = 0x087d
const XLS_Type_PAGELAYOUTVIEW = 0x088b
const XLS_Type_UNKNOWN = 0xffff

// Encryption type
const MS_BIFF_CRYPTO_NONE = 0
const MS_BIFF_CRYPTO_XOR = 1
const MS_BIFF_CRYPTO_RC4 = 2

// Size of stream blocks when using RC4 encryption
const REKEY_BLOCK = 0x400
