README 

EXCEL METADATA TO EML - Version 0.4  (updated 2021 December 1)

The standalone executables and the Perl script described below convert 
LTER EML Metadata Submission Template files (see xlsx2EML-04_Metadata_Template_FCE.xlsx) 
to EML 2.2.  The metadata template and this program are based on the 
EML Best Practices, Version 2 document released in August 2011. 

This program was developed with support from the Florida Coastal Everglades (FCE), 
Georgia Coastal Ecosystems (GCE), and Sevilleta (SEV) Long Term Ecological 
Research (LTER) programs.  Contributors to this program and the Excel metadata template 
include:

  Linda Powell, Kristin Vanderbilt, and Mike Rugge from Florida Coastal Everglades LTER Program 
  (http://fcelter.fiu.edu) at Florida International University.

  Wade Sheldon from Georgia Coastal Ecosystems LTER Program 
  (http://gce-lter.marsci.uga.edu/lter/) at the University of Georgia.

  Kristin Vanderbilt from the Sevilleta Long-Term Ecological Research LTER Program 
  (http://sevilleta.unm.edu) at the University of New Mexico.

  Youngmi Kim and Travis Brooks, programmers for the Canopy Database Project and graduates 
  of The Evergreen State College Software Engineering Program (http://canopy.evergreen.edu/).

  Judy Bayard Cushing, Ph.D., a member of the Faculty (Computer Science), The Evergreen State College, 
  Olympia, Washington and a principal investigator of the Canopy Database Project 
  (http://academic.evergreen.edu/j/judyc/home.htm, http://canopy.evergreen.edu/).

  Working in cooperation with the Evergreen State College contributors are Professor Barbara Bond, 
  Department of Forest Science, Oregon State University and her students, Georgianne Moore, 
  Texas A&M University and Kate George, USDA.
  
  
DOWNLOAD LOCATIONS

    GitHub - https://github.com/lter/FCE/tree/master/EML/SupportTool

    FCE LTER - http://fcelter.fiu.edu/research/information_management/tools/
      The FCE LTER download location includes links to the Windows zip files hosted
      on GitHub.


LICENSE

	This material is based upon work supported by National Science Foundation
	through the Florida Coastal Everglades Long-Term Ecological Research program
	under Cooperative Agreements #DEB-1832229, #DEB-1237517, #DBI-0620409, and #DEB-9910514. Any opinions,
	findings, conclusions, or recommendations expressed in the material are those
	of the author(s) and do not necessarily reflect the views of the National
	Science Foundation.
	
	Copyright (C) 2004, 2010, 2013, 2017, 2021  Florida International University
	
	This program is free software; you can redistribute it and/or
	modify it under the terms of the GNU General Public License
	as published by the Free Software Foundation; either version 2
	of the License, or (at your option) any later version.
	
	This program is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU General Public License for more details.
	
	You should have received a copy of the GNU General Public License
	along with this program; if not, write to the Free Software
	Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.


METADATA TEMPLATE

   xlsx2EML-04_Metadata_Template_FCE.xlsx - Metadata template

    Fill out the information in the five worksheets (General Metadata, MethodsCitation, MethodsProtol,
    ResearchProjects, and DataTable) according to the directions at the top of each worksheet and the 
    documentation in the found in a Microsoft Word help document called xlsx2EML-03_Metadata_Instructions.doc.  
    All unit definitions come from the 'Units IM Use Only' worksheet.  
    Additional custom units can be added to the bottom of the 'Units IM Use Only' worksheet.  


EXECUTABLE FILE

  xlsx2EML-04_Windows10.exe

    Converts the Metadata template above to an EML-compliant XML file.
    This version doesn't require Perl. 

    Windows 10 executables were generated from xlsx2EML-04_Tk.pl using Strawberry Perl 5.26 and the 
    PAR::Packer 1.037 Perl module (packages all of a script's required Perl components and modules into an 
    executable file).  

    Command used to generate executables for Windows:
      Windows 10:
	  pp -o xlsx2EML-03_Windows10.exe -l C:/Strawberry/c/bin/libcrypto-1_1-x64__.dll -l C:/Strawberry/c/bin/libiconv-2__.dll -l C:/Strawberry/c/bin/liblzma-5__.dll -l C:/Strawberry/c/bin/libssl-1_1-x64__.dll -l C:/Strawberry/c/bin/libxml2-2__.dll -l C:/Strawberry/c/bin/zlib1__.dll xlsx2EML-03_Tk.pl
      	        

PERL SCRIPT

    xlsx2EML-04_Tk.pl (Windows)

    Converts the Metadata template above to an EML-compliant XML file.
    Requires:
      - Perl 5.14 or higher

      - The following perl modules and their dependencies:
        - FindBin;
        - Config;
        - Spreadsheet::ParseExcel
        - use Spreadsheet::XLSX
        - use Spreadsheet::XLSX::Utility2007 qw(ExcelFmt ExcelLocaltime LocaltimeExcel)
        - OLE::Storage_Lite
        - IO::Scalar
        - Config
        - Cwd
        - XML::LibXML
        - IO::Handle
        - Getopt::Long

        - Digest::MD5::File qw(url_md5 url_md5_hex -utf8);
	- LWP::UserAgent;
	- LWP::Protocol::https;
	- IO::Socket::SSL;
        - HTTP::Request;      

    Platform-specific requirements:
      Windows (xlsx2EML-03_Tk.pl):
        - Tk
        - Tk::ProgressBar

    Change the shebang (first line of the script) to point to the location of your Perl installation.
    


NOTES

    - EML files will retain the same name as the Excel files, but their file extension will be 'xml' instead of 'xls' or 'xlsx'. 

    - The program will embed data entered in the Values section of the DataTable worksheet into the EML file if the 
    'Embed data' box is checked.  If no values are entered in this section, no data will be embedded in the EML. 
    If you plan to embed data, please use EML 2.1.0 or higher.

    - Some or all buttons may not be visible if your screen resolution is less than 1024 X 768.  If buttons aren't
    visible, try making the window larger or using the commands in the File menu at the top instead.

    - The program uses up a chunk of memory each time it converts an Excel file to EML.  This memory isn't returned to
    the system until after you exit the program.

       Tips for best performance:

       - Validate against local schema or deselect the validation option.  Validation against a schema URL can take 
       awhile if you have a slow network connection.

       - Keep the size of the Excel Metadata Template files as small as possible.  For example, if you don't have 
       any methods citations, you could clear all of the cells in the methodsCitation sheet to reduce the size of the file.

       Please note that every worksheet needs to be present in the specified order for the program to work, though.


INSTRUCTIONS

    1. Fill out the LTER EML Metadata Template (xlsx2EML-04_Metadata_Template_FCE.xlsx).  
    Instructions for filling out the template are provided in a Microsoft Word help document called xlsx2EML-04_Metadata_Instructions.doc.

    2. Add the template files to the list of files to convert to EML with the 'Add file to the list' button.  
    You can also use the 'Add file' in the file menu to add files to the list.  The other buttons and 
    choices in the file menu let you remove files, clear the log to the left, and exit the program.

    3. Fill out the Optional Conversion Information section.  
     Indentation
       The number of spaces per indent is optional.  If you leave this field blank, it defaults to two spaces 
       per indent. 

     EML Schema and STMML Schema
       The EML and STMML Schema URLs or file paths are optional and will default to eml.xsd and stmml.xsd, respectively.  
       However, in order to validate the final EML document, URLs or file paths that point to 
       the eml.xsd and stmml.xsd must be included (i.e. http://ltersite.edu/eml.xsd or C:\eml\eml.xsd). 
       EML schema validation is performed using the specified schemas after each EML file is created.  
       The stmml.xsd is only required if custom units are included in DataTable worksheet of the metadata template. 

     XSL Stylesheet
       The XSL Stylesheet URL is also optional and won't be included in the final EML files unless the URL is entered.  

     Embedded Data
       You can choose to embed data in the EML document by checking the box at the bottom of this section.  
       Please note that EML 2.1.0 or higher is required to validate EML files with embedded data.

     Validation 
       Validation against the specified EML schema (the EML and STMML Schema URLs or file paths) is optional, but 
       checked (enabled) by default. Validation warnings and errors will be displayed in the program's log and
       recorded in an error log file (error.log). 
    
    4. Select a destination directory for EML files.  
    If this field is blank, all EML files will be saved in the same directory as their source Excel file.  

    5. Click on the 'Convert all file to EML' button or select 'Convert to EML' from the file menu 
    to convert all files in the list to EML.  The log on the right displays messages if a file is successfully
    converted to EML (with the path to the new EML file), if a file isn't converted to EML, and if the file 
    has EML schema validation warnings or errors.  Validation warnings and errors are also recorded in an 
    error log file (error.log). Files in the list which don't have xls or xlsx extensions or the value 'Dataset Title' 
    in cell B20 will not be converted to EML.  Click on the 'Stop converting files' button or select 
    'Stop converting files' from the file menu if you wish to stop the conversion process.


FEEDBACK, BUG REPORTS

    Bug reports, comments, suggestions, and solutions to problems are always welcome. 

    Thank you

    Kristin Vanderbilt                     Mike Rugge
    krvander@fiu.edu                       ruggem@fiu.edu
    FCE LTER Information Manager           FCE LTER Program Manager

    

    



