{\rtf1\ansi\ansicpg1252\cocoartf2870
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 //Written by James Moss\
//September 10, 2016\
//\
// This script independently verifies and extends the distance calculations \
// in "share analysis TBD006625.xlsx" as part of a competitive effects \
// analysis for a chain of nine Valero-branded gas stations (TB II through \
// TB XVI and OB I) in the Phoenix, AZ metro area. The analysis was conducted \
// using OPIS gas station price and location data.\
//\
// The script replicates the great-circle distance calculations from the \
// Excel sheet in Stata, then constructs antitrust diversion ratios for each \
// focal store under three geographic market definitions: all stations, \
// stations within 10 miles, and stations within 5 miles. These metrics \
// measure the probability that a customer diverting away from a focal store \
// would choose another store in the same chain, and are used to assess \
// whether the focal stores face meaningful competitive constraints from \
// surrounding stations.\
//\
// Results are exported to "ShareAnalysis final.xlsx" with separate sheets \
// for station-level distances, competitor frequency counts, and diversion \
// ratio estimates.\
\
// \uc0\u9472 \u9472  SETUP \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \
clear all\
set more off\
\
// Open log file to record all output from this session\
capture log using "/Users/jbmoss/Documents/ShareAnalysis_Moss.log", replace\
\
// Load raw gas station data from Excel; first row contains variable names\
import excel "/Users/jbmoss/Downloads/share analysis TBD006625.xlsx", sheet("share analysis TBD006625") firstrow\
cd "/Users/jbmoss/Documents/"\
\
// \uc0\u9472 \u9472  DATA CLEANING \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \
// Drop observations with no station ID (unidentifiable records)\
drop if StationID==.\
\
// Drop variables not needed for this analysis\
drop Date TimeSummary State ZipCode OPISRegion MSA County Country CorporateBrand Product Price NetPrice WholesalePrice Margin\
\
// Create observation index for identifying the first record of each store\
gen obs=_n\
\
// \uc0\u9472 \u9472  DISTANCE CALCULATIONS \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \
// Define the list of focal TB/OB stores to compute distances from.\
// These are the stores whose geographic market areas are being analyzed.\
local TBStores `" "TB II" "TB III" "TB IV" "TB IX" "TB VI" "TB VII" "TB X" "TB XVI" "OB I" "'\
\
foreach store of local TBStores \{\
    display "`store'"\
    \
    // Get the observation index of the first record for this store,\
    // then extract its latitude and longitude as scalars for distance computation\
    sum obs if GasBrand=="`store'", meanonly\
    scalar StoreLat= Latitude[r(min)]\
    scalar StoreLong=Longitude[r(min)]\
    \
    // Remove spaces from store name to create a valid Stata variable name prefix\
    local storevar = subinstr("`store'"," ","",.) \
    \
    // Compute great-circle distance (in miles) from every station to this store\
    // using the Haversine-style acos formula; 3961 is Earth's radius in miles\
    gen `storevar'distance=acos(sin(Latitude*_pi/180)*sin(StoreLat*_pi/180)+cos(Latitude*_pi/180)*cos(StoreLat*_pi/180)*cos((StoreLong-Longitude)*_pi/180))*3961\
    \
    // Create distance variables restricted to stations within 10 and 5 miles\
    // (missing for stations outside the radius \'97 used in market share calculations)\
    gen `storevar'distance10=`storevar'distance if `storevar'distance<10\
    gen `storevar'distance5=`storevar'distance if `storevar'distance<5\
\}\
\
// Save dataset with all distance variables before collapsing to brand level\
save "ShareAnalysisDistancesTBD006625", replace\
\
// \uc0\u9472 \u9472  FREQUENCY COUNTS \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \
// Count the number of stations of each GasBrand within each distance radius\
// for every focal store. bysort+egen count gives brand-level totals.\
local TBStores3 TBII TBIII TBIV TBIX TBVI TBVII TBX TBXVI OBI\
local distancevars TB* OB*\
foreach var of varlist `distancevars' \{\
    bysort GasBrand: egen Freq`var'= count(`var')\
\}\
\
// Collapse to one row per GasBrand (all Freq* values are constant within brand)\
keep GasBrand Freq*\
bysort GasBrand: drop if _n!=1\
\
// Rename Freq variables: drop "distance" from names for cleaner output\
// e.g. FreqTBIIdistance10 \uc0\u8594  FreqTBII10\
local CountVars Freq*\
foreach var of varlist `CountVars'\{\
    local newvarname = subinstr("`var'", "distance", "", 1)\
    rename `var' `newvarname' \
\}\
\
// \uc0\u9472 \u9472  MARKET SHARE CHANGE TABLES \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \
// For each focal store and each distance band (all, <10mi, <5mi), compute\
// two market share diversion metrics:\
//\
//   MSChangeTB: share diverted to/from TB/OB stores only\
//     = (brand count in radius) / (total count * (total count - 1))\
//\
//   MSChangeAll: share diverted relative to all TB/OB stores combined\
//     = (brand count * total TB/OB count) / (total * (total - TB/OB count))\
//     Adjusts for the combined presence of all TB/OB stores in the radius.\
\
local TBStores2 TBII TBIII TBIV TBIX TBVI TBVII TBX TBXVI OBI\
foreach store of local TBStores2 \{\
    foreach distance in "" "10" "5" \{\
    \
        // Total number of stations of any brand within this radius\
        egen TotalCount=total(Freq`store'`distance') \
        gen MSChangeTB`store'`distance'= Freq`store'`distance'/(TotalCount*(TotalCount-1))\
        \
        // Count of TB/OB stores specifically within this radius\
        egen y=total(Freq`store'`distance') if inlist(GasBrand, "TB II", "TB III", "TB IV", "TB IX", "TB VI", "TB VII", "TB X", "TB XVI", "OB I")\
        egen TBCount= mean(y)\
        \
        gen MSChangeAll`store'`distance'=(Freq`store'`distance'*TBCount)/(TotalCount*(TotalCount-TBCount))\
        \
        // Drop temporary variables before next iteration\
        drop TotalCount\
        drop y\
        drop TBCount\
    \}\
\}\
\
// \uc0\u9472 \u9472  EXPORT \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \u9472 \
save "ShareAnalysis tabs share.dta", replace\
\
// Export frequency count tables (all brands, all radius sizes)\
export excel GasBrand FreqTB*10 FreqOB*10 FreqTB*5 FreqOB*5 FreqTB*V FreqTB*I FreqTB*X FreqOBI using "ShareAnalysis final.xlsx", sheet("Frequency Tables") sheetreplace firstrow(variables)\
\
// Restrict to non-TB/OB brands for market share export\
// (TB/OB stores are the focal stores, not the competitors being measured)\
drop if inlist(GasBrand, "TB II", "TB III", "TB IV", "TB IX", "TB VI", "TB VII", "TB X", "TB XVI", "OB I")\
\
// Export market share diversion relative to TB/OB stores only\
export excel GasBrand MSChangeTBTB*10 MSChangeTBOB*10 MSChangeTBTB*5 MSChangeTBOB*5 MSChangeTBTB*V MSChangeTBTB*I MSChangeTBTB*X MSChangeTBOBI using "ShareAnalysis final.xlsx", sheet("MarketShareTB") sheetreplace firstrow(variables)\
\
// Export market share diversion relative to all stations in radius\
export excel GasBrand MSChangeAllTB*10 MSChangeAllOB*10 MSChangeAllTB*5 MSChangeAllOB*5 MSChangeAllTB*V MSChangeAllTB*I MSChangeAllTB*X MSChangeAllOBI using "ShareAnalysis final.xlsx", sheet("MarketShareAll") sheetreplace firstrow(variables)\
\
// Reload full distance dataset and export raw distances for all stations\
use "ShareAnalysisDistancesTBD006625.dta", clear\
export excel StationID StationName Address City GasBrand Latitude Longitude StoreBrand TB*distance10 OB*distance10 TB*distance5 OB*distance5 TB*distance OB*distance using "ShareAnalysis final.xlsx", sheet("Distances") sheetreplace firstrow(variables)\
\
log close}