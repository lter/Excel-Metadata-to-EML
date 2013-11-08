README 

EXCEL METADATA TO EML - Version 0.2  (updated 11/1/2013)

The standalone executables and the Perl script described below convert 
LTER EML Metadata Submission Template files (see xlsx_eml_02.xls) 
to EML 2.1.  The metadata template and this program are based on the 
EML Best Practices, Version 2 document released in August 2011. 

This program was developed with support from the Florida Coastal Everglades (FCE), 
Georgia Coastal Ecosystems (GCE), and Sevilleta (SEV) Long Term Ecological 
Research (LTER) programs.  Contributors to this program and the Excel metadata template 
include:

  Linda Powell and Mike Rugge from Florida Coastal Everglades LTER Program 
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

    SVN - https://svn.lternet.edu/websvn/listing.php?repname=FCE&path=/trunk/EML/SupportTool/&#a7832dec0373057bfba77dab17c0d1b08
      The SVN is hosted by the LTER Network Office and includes the latest version of all of the 
      files mentioned below except for the executable files. 

    FCE LTER - http://fcelter.fiu.edu/research/information_management/tools/
      The FCE LTER download location includes links to the Windows and Mac OS X zip files hosted
      on the LTER Network Office SVN.


LICENSE

	This material is based upon work supported by National Science Foundation
	through the Florida Coastal Everglades Long-Term Ecological Research program
	under Cooperative Agreements #DEB-1237517, #DBI-0620409, and #DEB-9910514. Any opinions,
	findings, conclusions, or recommendations expressed in the material are those
	of the author(s) and do not necessarily reflect the views of the National
	Science Foundation.
	
	Copyright (C) 2004, 2010, 2013  Florida International University
	
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

   xlsx2EML-02_Metadata_Template_FCE.xlsx - Metadata template

    Fill out the information in the five worksheets (General Metadata, MethodsCitation, MethodsProtol,
    ResearchProjects, and DataTable) according to the directions at the top of each worksheet and the 
    documentation in the found in a Microsoft Word help document called xlsx2EML-02_Metadata_Instructions.doc.  
    All unit definitions come from the 'Units IM Use Only' worksheet.  
    Additional custom units can be added to the bottom of the 'Units IM Use Only' worksheet.  


EXECUTABLE FILE

  xlsx2EML-02.exe (Windows) or xlsx2EML-02-MacOSX.dmg (Mac OS X) - Executable file

    Converts the Metadata template above to an EML-compliant XML file.
    This version doesn't require Perl. 

    The Windows executable was generated on Windows XP from xlsx2EML-02_Tk.pl using Perl 5.14 and the 
    PAR::Packer 1.014 Perl module (packages all of a script's required Perl components and modules into an 
    executable file).  The Mac OS X executable was generated on Mac OS X Snow Leopard (10.6.8) from 
    xlsx2EML-02_Tcl.pl using Citrus Perl 5.16.1 (http://www.citrusperl.com/) and Cava Packager (http://www.cavapackager.com/), 
    which can create Mac OS X applications from Perl scripts.

    Command used to generate executable:
      Windows:
        pp --icon="icon3.ico" -o xlsx2EML-02.exe xlsx2EML-02_Tk.pl
      Mac OS X:
        Use Cava Packager to create an application bundle
	        1. Create a new project
	        2. Project->Project Details->Project Name: xlsx2EML-02
	        3. Project->Perl Interpreter: Use Citrus Perl for the Perl Interpreter.  
	        4. Project->Perl Interpreter: Add perl/lib and perl/site/lib to "Extra Module Search Paths" under Perl Interpreter.
	        5. Executables: Use xlsx2EML-02_Tcl.pl for the "Packaged Script"
	        6. Scripts: Use xlsx2EML-02_Tcl.p
	        7. Force Include Modules: Include the following Perl modules
					Class::ISA
					Find::Bin
					Getopt::Long
					IO::Handle
					Spreadsheet::ParseExcel
					Spreadsheet::XLSX
					Sub::Name
					Tcl::pTk
					XML::LibXML
					XML::LibXML::Reader
					deprecate
	        8. Shared Libraries: libperl.dylib (perl/lib/CORE/libperl.dylib)
	        9. Scan and build the project
	        10. The release version (xlsx2EML-02.app) will be in the release directory of the directory specified for the project
	        11. The installer version (xlsx2EML-02.dmg) will be in the installer directory of the directory specified for the project
	        

PERL SCRIPT

    xlsx2EML-02_Tk.pl (Windows) or xlsx2EML-02_Tcl.pl (Mac OS X) - Perl scripts used to create the executables above

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
        - XML::LibXML::Reader
		- XML::LibXML
        - IO::Handle
        - Getopt::Long

    Platform-specific requirements:
      Windows (xlsx2EML-02_Tk.pl):
        - Tk
        - Tk::ProgressBar

      Mac OS X (xlsx2EML-02_Tcl.pl):
        - Tcl::pTk (qw/ :perlTk/)
        - Tcl::pTk::ProgressBar
        - Tcl::pTk::Tile

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

    1. Fill out the LTER EML Metadata Template (xlsx2EML-02_Metadata_Template_FCE.xlsx).  
    Instructions for filling out the template are provided in a Microsoft Word help document called xlsx2EML-02_Metadata_Instructions.doc.

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

    Linda Powell                           Mike Rugge
    powell@fiu.edu                         ruggem@fiu.edu
    FCE LTER Information Manager           FCE LTER Program Manager

    

    



