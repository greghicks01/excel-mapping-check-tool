# Excel Map to Source Tool

This tool is based on some work i have done for a commercial data engine.

A dicussion revealed soruce data headings have been imported with inappropriate characters and lengths. This tool is intended to provide a frontloading step to manage the bad characters for the headers and identify data that can't be loaded.

## Motivation

I developed this tool in my spare time, which gave the team a faster method to review and correct the input data than wait till an upload fails

## Journal

This table is upadted for each push and shows insights to the discussion, analysis and specification that we are building in to the tool

|| Journal Id || Date|| Discussion|| Feature to add || Push || Date completed ||
| 1|  23 Jul 22 | Source Data headers contain incorrect characters or are too long |  | create sheet scanning codebase | |
| | | | | | 
| | | | | | 
| | | | | | 
| | | | | | 

## Full Journal

### Journal Id 1.
#### Source Data headers contain incorrect characters or are too long
##### Discussion
External source reports under go updates beyond the control of the team that analyses the data supplied. 
This has lead to characters and lengths that breach the analysis tools input constraints.
The upload process fails under these conditions requiring manual fixes before upload can re-commence

##### Analysis
A tool that helps identify and potentially fix issues/discrepancies can reduce upload failures.

The bad characters are detectable in several ways:
a. as a character in a string
b. as a regex in a string (Regex.test)
The excess length is detectable as gt or lt

This suggests a set of filters need to be used for visual promting:
No Filter
Bad Characater Filter
Too Long Filter
Bad Char and Too long

A means of configurating rather than hard coding mean users can apply this to a broader set of problems.

A means of reading until a condition is met in either row-wise or columnwise
A means of reading until a condition is met in a control column or row and either rowwise or column wise

##### Constraints
excel sheets start Row=1 Column=1

##### Assumptions
Initial code will assume source data starts at Row=1 Column=1
#### Features identified
1. User configuraiton component reads from excel sheet
2. Open a mapping file, read the mapping names, display in a listbox
3. open the source file, get a list of sheets and a list of headers (assume cells start Row=1, Col = 1)

