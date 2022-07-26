# Excel Map to Source Tool

This tool is based on some work i have done for a commercial project.

A dicussion revealed source data headings have been imported with inappropriate characters and lengths.

## Motivation

I developed this tool in my spare time, which gave the team a faster method to review and correct the input data than wait till an upload fails

## Journal

This table is updated for each push and shows insights to the discussion, analysis and specification that we are building in to the tool

| Journal Id    | Date      | Discussion |  Date completed |
| :-----------: | :-------: | ---        |  ---            | 
| 1             | 23 Jul 22 | Source Data headers contain incorrect characters or are too long | 25 jul 22 |
| 2             | 25 Jul 22 | Scan rows for control condition | 25 Jul 22 |
| 3             | 26 jul 22 | File Picker Dialog with options | 26 Jul 22 |

## Full Journal

### Journal Id 1. Source Data headers contain incorrect characters or are too long
#### Discussion
External source reports under go updates beyond the control of the team that analyses the data supplied. 
This has lead to characters and lengths that breach the analysis tools input constraints.
The upload process fails under these conditions requiring manual fixes before upload can re-commence

#### Analysis
A tool that helps identify and potentially fix issues/discrepancies can reduce upload failures.

The bad characters are detectable in several ways:
a. as a character in a string
b. as a regex in a string (Regex.test)
The excess length is detectable as gt or lt

This suggests a set of filters need to be used for visual promting:
* no Filter
* Bad Characater Filter
* Too Long Filter
* Bad Char and Too long

#### Solution outline
To implement the above, there needs to be a means of 
* configurating rather than hard coding mean users can apply this to a broader set of problems.
* reading until a condition is met in either row-wise or columnwise
* reading until a condition is met in a control column or row and either rowwise or column wise

#### Constraints
excel sheets start Row=1 Column=1

#### Assumptions
excel sheets start Row=1 Column=1

#### Features identified
1. Open a mapping file, read the mapping names, display in a listbox
2. open the source file, get a list of sheets and a list of headers per sheet (assume cells start Row=1, Col = 1)
3. User configuraiton component reads from excel sheet

#### Dev Notes
All functions tested using a local worksheet. 
Caveat - remote connection/dwonload is not considered in the testing


## Journal Id 2. Scan rows for control condition
#### Discussion
Next issue is to run down a control column and collect all the key column data as a collection, followed by removing any blanks (compacting)

#### Analysis
must hand in multiple values and have similar approach to the previous work.

#### Solution outline
while loop <> key value
	if cell(row, controlcol) == stop then exit
	increment row
wend

#### Constraints
excel sheets start Row=1 Column=1

#### Assumptions
excel sheets start Row=1 Column=1 as defaults

#### Features identified
scan rows until key or control stop condition

#### Dev Notes
a little tech debt in comments to fox
identified a refactoring of the range checking iif into a consolidated function (maintenance)


## Journal Id 3. File Picker with default XL and CSV options
#### Discussion
The user will need to pick a either a single mapping file or a set of data files

#### Analysis
use the built in dialog and make most items available as parameters

#### Solution outline
wrap the dialog in s function and pass params to make it flexable
ensure that one data type is returned, preferrably iterator type to catch cancel condition (ie empty set)

#### Constraints
excel sheets start Row=1 Column=1

#### Assumptions
excel sheets start Row=1 Column=1 as defaults

#### Features identified
Give the dialog a title to give a user a clue what is to be picked
drive single or mulitple selects to ease picking the files

#### Dev Notes
Feature added and unit tested with all passes
Refacting improved the maintenance on the range checking