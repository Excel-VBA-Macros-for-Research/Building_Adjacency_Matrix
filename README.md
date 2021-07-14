# Excel VBA macros for building adjacency matrix
Adjacency matrix is ....   
The Excel VBA macros for building adjacency matrix on this site can make adjacency matrix almost automatically from Excel files of data about grouping. Although the VBA macros are optimized for Windows, it can also be used on Mac.  

## "Material_Sample" file
The "Material_Sample" file represents a data layout which can be used for building adjacency matrix with the VBA macros.
- The notation for ID and groups can be either numbers or letters.  
- The initial character of ID and the name of groups must not be "0" (zero).  
- There must be no duplicate in the "ID" column.  
- The notations for the same group must match exactly. Lowercase and uppercase letters are distinguished.  
- If a person belongs to no group, the cell must be blank. In such a case, no symbols, letters, or numbers must be entered in the cell.  
- The "Group_Type1" column represents a pattern in which each person belongs to a single group.  
- The "Group_Type2" column represents a pattern in which some persons belong to a group and the others do not belong to the group.  
- The "Group_Type3" column represents a pattern in which some persons belong to multiple groups. Each notation for multiple groups must be connected by semicolons in cells. Spaces must not be placed before or after semicolons.

## "Macros" file
The "Macros" file contains three macro buttons. Open an Excel file of data about grouping like "Material_Sample" file before executing the macros.
### Conversion_from_Individuals_to_Groups  
- Click the "Conversion_from_Individuals_to_Groups" button. Then, a dialog box appears.  
- Select the Excel file of data about grouping.  
- Click a cell in the column of intended category and click the OK button. Any cell in the intended column will do.  
- On Mac, another dialog box will appear. Then, click the "OK" button.  
- An Excel file displaying the members of each group is created in seconds. As for persons who belong to no groups, the groups whose names are IDs of the persons are placed for the sake of convenience. It is recommended that you save this file. The file should not be closed yet.  
### Building_Adjacency_Matrix_from_Groups  
- Click the "Building_Adjacency_Matrix_from_Group" button. Then, a dialog box appears.
- Select the Excel file which have been just created by "Conversion_from_Individuals_to_Groups" macro. If you already have an Excel file whose layout is similar to such a file displaying the members of each group, you can skip the previous step of "Conversion_from_Individuals_to_Groups" and start with this step.
- On Mac, another dialog box will appear. Then, click the "OK" button.  
- An Excel file showing adjacency matrix with "0" and "1" is created. It might take a few tens of seconds to a minute.  
### Recover from errors  
- Some errors might occur in executing macros when the layout of the material data file is not suited, for example.  
- If you encounter errors and executing macros are aborted, click the "Recover from errors" button in the "Macros" file before anything else calmly. Then, close all the files without saving them. Until the "Recover from errors" button is clicked, the display of any Excel files may be wrong.  

## "Source_Code" file
The "Source_Code_Version1" is a BAS file containing the source code of the Excel VBA macros. You can import this file in Excel.  

## About  
- Author: Mitsuyuki Numasawa  
- Affiliation: Institute of Education, Tokyo Medical and Dental University (TMDU), Tokyo, Japan.  
