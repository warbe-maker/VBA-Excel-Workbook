## VBA Excel Workbook Common Services
### Summary of Services

Most of the services accept a Workbook argument which is either a Workbook object, a Workbook's _FullName_ or a Workbook's _Name_. All services provide a comprehensive inline documentation. 

| Name         | Service |
| ------------ | ---------------------------------------------------------------------------------- |
| _Value_      | Let: Writes a value to a named range<br>Get: Returns a value from a named range  |
| _WbClose_    | Closes a Workbook if it is open.        |
| _Exists_     | Universal existence check for Workbook, worksheet, and Range Name. An existing and open Workbook is returned as Workbook object, an existing Worksheet is returned as Worksheet object |
| _GetOpen_    | Opens a Workbook and returns it as Workbook object       |
| _IsFullName_ | Returns TRUE when a provided string represents a Workbook's _FullName_ whereby the Workbook must neither be open nor existent.  |
| _isName_     | Returns TRUE when a provided string represents a Workbook's _Name_       |
| _IsWbObject_ | Returns TRUE when a provided argument represents a Workbook object       |
| _IsWsObject_ | Returns TRUE when a provided argument represents a Workbooks Worksheet either by its _Name_ or _by its _CodeName        |
| _IsOpen_     | Returns TRUE when a provided argument identifies an open Workbook whereby the open Workbook is retrurned as Workbook object        |
| _Opened_     | Returns a Dictionary with all open Workbooks in all application instances. |


### Installation
Download and import _[mWbk.bas][1]_

### Usage
> This _Common Component_ is prepared to function completely autonomously ( download, import, use) but at the same time to integrate with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][2] for more details.


## Contribution
Any kind of contribution is absolutely and unconditionally welcome. 

[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Workbook-Services/master/source/mWbk.bas
[2]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html