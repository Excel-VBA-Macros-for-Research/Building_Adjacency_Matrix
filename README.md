# Excel VBA for building adjacency matrix
Adjacency matrix is .... The excel VBA for building adjacency matrix on this site can make adjacency matrix almost automatically from excel files of data about grouping. Although the VBA are optimized for Windows, it can also be used in Mac.

## Material_Sample.xlsx  
The "Material_Sample" file represents a data layout which can be used for building adjacency matrix with the VBA.
- The notation for ID and groups can be either numbers or letters.  
- The initial character of ID and the name of groups must not be zero.  
- There must be no duplicate in the "ID" column.  
- The notations for the same group must match exactly. Lowercase and uppercase letters are distinguished.
- If a person belongs to no group, the cell must be blank. In such a case, no symbols, letters, or numbers must be entered in the cell.
- The "Group_Type1" column represents a pattern in which each person belongs to a single group.
- The "Group_Type2" column represents a pattern in which some persons belong to a group and the others do not belong to the group.
- The "Group_Type3" column represents a pattern in which some persons belong to multiple groups. Each notation for multiple groups must be connected by a semicolon in cells. Spaces must not be placed before or after semicolons.
## XLSM_file_Version1.xlsm  
The "XLSM file" 
