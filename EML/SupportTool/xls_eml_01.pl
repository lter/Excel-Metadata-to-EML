#!c:/perl/bin/perl
use strict;
use Config;
use Spreadsheet::ParseExcel;
use Tk;
use Tk::Scrollbar;
use Tk::ProgressBar;
use Tk::DialogBox;
use Cwd;
use XML::Xerces;
use IO::Handle;
use Getopt::Long;

################################################################################
#
# Excel Metadata to EML - Version 0.1
#
# This program converts LTER EML Metadata Submission Template files
# (in Excel format) to EML 2.0.1 files.
#
# This material is based upon work supported by the National Science Foundation
# under Cooperative Agreement #DEB-9910514. Any opinions, findings, conclusions,
# or recommendations expressed in the material are those of the author(s) and
# do not necessarily reflect the views of the National Science Foundation.
#
# Copyright (C) 2004  Florida International University
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
#
################################################################################

##########################
#  Frames for GUI  (Tk)  #
##########################
my $mw;
my @files;
my $loop;

$mw = MainWindow->new( -background => '#0C3380' );
$mw->configure( -menu => my $menu = $mw->Menu );
$mw->title("Excel Metadata to EML");

my $top = $mw->Frame( -background => '#0C3380' )->pack(
    -side => 'top',
    -fill => 'x'
);

my $top_content = $mw->Frame( -background => '#0C3380' )->pack(
    -side => 'top',
    -fill => 'x',
    -padx => 10,
    -pady => 10
);

my $top_text = $top_content->Frame( -background => '#FFFFFF' )->pack(
    -side  => 'right',
    -fill  => 'x',
    -ipadx => 5,
    -ipady => 5,
);

my $top_entries = $top_content->Frame( -background => '#FFFFFF' )->pack(
    -side  => 'left',
    -fill  => 'x',
    -ipadx => 5,
    -ipady => 5
);

my $top_entries1 = $top_entries->Frame( -background => '#FFFFFF' )->pack(
    -side => 'top',
    -fill => 'x',
);

my $top_entries2 = $top_entries->Frame( -background => '#FFFFFF' )->pack(
    -side => 'top',
    -fill => 'x'
);

my $top_entries3 = $top_entries->Frame( -background => '#FFFFFF' )->pack(
    -side => 'top',
    -fill => 'x'
);

my $top_entries4 = $top_entries->Frame( -background => '#FFFFFF' )->pack(
    -side => 'top',
    -fill => 'x'
);

my $top_entries5 = $top_entries->Frame( -background => '#FFFFFF' )->pack(
    -side => 'top',
    -fill => 'x'
);

my $top_entries6 = $top_entries->Frame( -background => '#FFFFFF' )->pack(
    -side => 'top',
    -fill => 'x'
);

my $top_entries7 = $top_entries->Frame( -background => '#FFFFFF' )->pack(
    -side => 'top',
    -fill => 'x'
);

my $bottom = $mw->Frame( -background => '#0C3380' )->pack(
    -side => 'bottom',
    -fill => 'x'
);

my $lb_left = $mw->Frame( -background => '#0C3380' )->pack(
    -side => 'left',
    -fill => 'x',
    -padx => 5,
    -pady => 5
);

my $lb_left_top = $lb_left->Frame( -background => '#0C3380' )->pack(
    -side   => 'top',
    -fill   => 'both',
    -expand => 1,
    -padx   => 5,
    -pady   => 5
);

my $lb_right = $mw->Frame( -background => '#0C3380' )->pack(
    -side => 'right',
    -fill => 'x',
    -padx => 5,
    -pady => 5
);

my $file_menu = $menu->cascade( -label => "~File", -tearoff => 0 );
my $help_menu = $menu->cascade( -label => "~Help", -tearoff => 0 );
$file_menu->command( -label => "~Add file",                -command => \&getFile );
$file_menu->command( -label => "~Convert to EML",          -command => \&convertToEML );
$file_menu->command( -label => "~Remove selected file(s)", -command => \&removeFile );
$file_menu->command( -label => "~Remove all files",        -command => \&removeAllFiles );
$file_menu->command( -label => "~Clear log",               -command => \&clearLog );
$file_menu->command( -label => "~Stop converting files",   -command => sub { $loop = 0 } );
$file_menu->command( -label => "~Exit", -command => sub { exit; } );
$help_menu->command( -label => "~Instructions",       -command => \&getInstructions );
$help_menu->command( -label => "~Notes",              -command => \&getNotes );
$help_menu->command( -label => "~About this program", -command => \&getInfo );

###################################################
#  Listboxes for the file list and the log  (Tk)  #
###################################################

my $lb_in = $lb_left->Scrolled(
    "Listbox",
    -scrollbars => "osoe",
    -width      => "70",
    -height     => "20",
    -background => 'white',
    -relief     => 'sunken',
    -selectmode => 'multiple'
)->pack( -side => "bottom" );

my $lb_out = $lb_right->Scrolled(
    "Listbox",
    -scrollbars => "osoe",
    -width      => "70",
    -height     => "23",
    -background => 'white',
    -relief     => 'sunken',
    -selectmode => 'multiple'
)->pack( -side => "bottom" );

####################################################################
#  Label and entry widgets for optional information section  (Tk)  #
####################################################################
my $OS = "$^O";
my $basefontsize;
my $fontfamily;

if ( $OS =~ /mswin/i ) {
    $basefontsize = "10";
    $fontfamily   = "Arial, Helvetica";
}
else {
    $basefontsize = "12";
    $fontfamily   = "Arial, Helvetica, Utopia";
}

$top->Label(
    -font       => [ -size => $basefontsize + 6, -family => $fontfamily, -weight => 'bold' ],
    -text       => 'Excel Metadata to EML - Version 0.1',
    -foreground => '#FFFF99',
    -background => '#0C3380',
    -anchor     => 'n'
)->pack( -side => "top" );

$top_entries1->Label(
    -font       => [ -size => $basefontsize, -family => $fontfamily, -weight => 'bold' ],
    -text       => 'Optional conversion information',
    -foreground => 'black',
    -background => '#FFFFFF',
)->pack( -side => "left" );

my $indent_level;
$top_entries2->Label(
    -font       => [ -size => $basefontsize, -family => $fontfamily, -weight => 'normal' ],
    -text       => 'Number of spaces per indent (default=2)     ',
    -foreground => 'black',
    -background => '#FFFFFF',
)->pack( -side => "left" );

$top_entries2->Entry(
    -width        => '10',
    -textvariable => \$indent_level,
)->pack( -side => "left" );

$top_entries3->Label(
    -font       => [ -size => $basefontsize, -family => $fontfamily, -weight => 'normal' ],
    -text       => 'EML Schema URL (defaults to LNO URL)  ',
    -foreground => 'black',
    -background => '#FFFFFF',
)->pack( -side => "left" );

my $schema;
$top_entries3->Entry(
    -width        => '28',
    -textvariable => \$schema,
)->pack( -side => "left" );

$top_entries4->Label(
    -font       => [ -size => $basefontsize, -family => $fontfamily, -weight => 'normal' ],
    -text       => 'STMML Schema URL (defaults to LNO URL)  ',
    -foreground => 'black',
    -background => '#FFFFFF',
)->pack( -side => "left" );

my $stmml;
$top_entries4->Entry(
    -width        => '28',
    -textvariable => \$stmml,
)->pack( -side => "left" );

$top_entries5->Label(
    -font       => [ -size => $basefontsize, -family => $fontfamily, -weight => 'normal' ],
    -text       => 'XSL Stylesheet URL (optional)  ',
    -foreground => 'black',
    -background => '#FFFFFF',
)->pack( -side => "left" );

my $stylesheet;
$top_entries5->Entry(
    -width        => '28',
    -textvariable => \$stylesheet,
)->pack( -side => "left" );

my $embed_data_checkbox;
$top_entries6->Checkbutton(
    -font             => [ -size => $basefontsize, -family => $fontfamily, -weight => 'normal' ],
    -text             => 'Embed data in EML (EML 2.0.1 and higher only)',
    -onvalue          => 'yes',
    -offvalue         => 'no',
    -background       => '#FFFFFF',
    -activebackground => '#FFFFFF',
    -variable         => \$embed_data_checkbox
)->pack( -side => "left" );

my $validation_checkbox = "yes";
$top_entries7->Checkbutton(
    -font             => [ -size => $basefontsize, -family => $fontfamily, -weight => 'normal' ],
    -text             => 'Validate against the specified EML schema',
    -onvalue          => 'yes',
    -offvalue         => 'no',
    -background       => '#FFFFFF',
    -activebackground => '#FFFFFF',
    -variable         => \$validation_checkbox
)->pack( -side => "left" );

############################################
#  Label  for instructions  section  (Tk)  #
############################################

$top_text->Label(
    -font => [ -size => $basefontsize, -family => $fontfamily, -weight => 'normal' ],
    -text =>
'This program converts LTER EML Metadata Submission Template files (in Excel format) to EML 2.0.1 files. Add files to the list on the left and press the \'Convert all files to EML\' button to convert files to EML.

The EML files will retain the same name as the Excel files, but their file extension will be xml instead of xls.  The Metadata Submission Template files must have a xls extension in order to be converted to EML.',
    -foreground => 'black',
    -background => '#FFFFFF',

    #    -width      => '50',
    #    -height     => '8',
    -justify    => 'left',
    -wraplength => '400'
)->pack( -side => "left" );

######################################################################
#  Label and button widgets for the EML File List and the Log  (Tk)  #
######################################################################

$lb_left_top->Label(
    -font       => [ -size => $basefontsize + 2, -family => $fontfamily, -weight => 'bold' ],
    -text       => 'List of Excel Metadata files to convert to EML',
    -foreground => '#FFFF99',
    -background => '#0C3380',
    -anchor     => 'n'
)->pack( -side => "top" );

if ( &perl_ver >= 58 ) {
    $lb_left_top->Button(
        -text             => 'Destination directory for EML files',
        -command          => \&chooseDir,
        -background       => '#F3F379',
        -activebackground => '#FFFF99'
    )->pack( -side => "left" );
}
else {
    $lb_left_top->Label(
        -font       => [ -size => $basefontsize, -family => $fontfamily, -weight => 'normal' ],
        -text       => 'Destination directory for EML files',
        -foreground => '#FFFF99',
        -background => '#0C3380',
        -anchor     => 'n'
    )->pack( -side => "left" );

}

my $save_dir;
my $dir_entry = $lb_left_top->Entry(
    -width        => '40',
    -textvariable => \$save_dir,
)->pack( -side => "left" );

$lb_left->Button(
    -text             => 'Add file to the list',
    -command          => \&getFile,
    -background       => '#F3F379',
    -activebackground => '#FFFF99'
)->pack( -side => "left" );

$lb_left->Button(
    -text             => 'Remove selected file(s)',
    -command          => \&removeFile,
    -background       => '#F3F379',
    -activebackground => '#FFFF99'
)->pack( -side => "left" );

$lb_left->Button(
    -text             => 'Remove all files',
    -command          => \&removeAllFiles,
    -background       => '#F3F379',
    -activebackground => '#FFFF99'
)->pack( -side => "left" );

$lb_right->Label(
    -font       => [ -size => $basefontsize + 2, -family => $fontfamily, -weight => 'bold' ],
    -text       => 'Log',
    -foreground => '#FFFF99',
    -background => '#0C3380',
    -anchor     => 'n'
)->pack( -side => "left" );

$lb_right->Button(
    -text             => 'Clear log',
    -command          => \&clearLog,
    -background       => '#F3F379',
    -activebackground => '#FFFF99'
)->pack( -side => "right", -padx => 5 );

$lb_right->Button(
    -text             => 'Stop converting files',
    -command          => sub { $loop = 0 },
    -background       => '#D95757',
    -activebackground => '#F36161'
)->pack( -side => "right", -padx => 10 );

$lb_left->Button(
    -text             => 'Convert all files to EML',
    -command          => \&convertToEML,
    -background       => '#6CD998',
    -activebackground => '#79F3AA'
)->pack( -side => "right" );

####################################################################################
#  Label, progress bar, and exit button widgets at the bottom of the window  (Tk)  #
####################################################################################

my $progress;
my $percent_done;

$bottom->Button(
    -text             => "Exit",
    -command          => sub { exit; },
    -background       => '#F3F379',
    -activebackground => '#FFFF99'
)->pack( -side => "right" );

$bottom->Label(
    -font       => [ -size => $basefontsize + 2, -family => $fontfamily, -weight => 'bold' ],
    -text       => '           ',
    -foreground => '#FFFF99',
    -background => '#0C3380'
)->pack( -side => "right" );

$progress = $bottom->ProgressBar(
    -width    => 20,
    -length   => 250,
    -from     => 0,
    -to       => 100,
    -blocks   => 50,
    -colors   => [ 0, '#6CD998', 50, '#6CD998', 100, '#6CD998' ],
#    -variable => \$percent_done
)->pack( -side => "right" );

$bottom->Label(
    -font       => [ -size => 12, -weight => 'bold' ],
    -text       => 'Progress ',
    -foreground => '#FFFF99',
    -background => '#0C3380'
)->pack( -side => "right" );

MainLoop;

####################
#  Tk Subroutines  #
####################

# Detect Perl version (check to see if it's >5.8)
sub perl_ver {
    my %config;
    my $perl_version = $Config{version};
    my @perl_version = split( /\./, $perl_version );
    my $perl         = "$perl_version[0]" . "$perl_version[1]";
    return $perl;
}

# 'Destination directory for EML files'  action
sub chooseDir {

    my $save_dir;
    $dir_entry->delete( 0, 500 );
    $save_dir = $mw->chooseDirectory( -parent => $mw );
    $dir_entry->insert( 0, $save_dir );

}

# 'Add file to the list' action
sub getFile {
    my $OS = "$^O";
    my $mult;

    if ( $OS =~ /mswin/i ) {
        $mult = "1";
    }
    else {
        $mult = "0";
    }
    my $f;
    my @file = ();

    # Types are listed in the dialog widget
    my $types = [ [ "Excel Metadata template", '.xls' ], [ "All Files", "*" ] ];

    if ( &perl_ver >= 58 ) {
        @file = $mw->getOpenFile( -filetypes => $types, -multiple => $mult );
    }
    else {
        @file = $mw->getOpenFile( -filetypes => $types );
    }

    foreach $f (@file) {
        if ( $f gt '' ) {
            $lb_in->insert( "end", $f );
        }
    }

}

# 'Remove selected file(s)' action
sub removeFile {

    my @selected = ( $lb_in->curselection );

    my $item;
    foreach $item (@selected) {
        my $file_deleted = $lb_in->get($item);
        $lb_in->delete($item);
    }

}

# 'Remove all files' action
sub removeAllFiles {

    my $listboxsize = $lb_in->size;
    my $size        = $listboxsize - 1;

    $lb_in->delete( 0, $size );

    #    my $deleted = "Removed all files";
    #    $lb_out->insert( "end", $deleted );
}

# 'Clear log'  action
sub clearLog {

    my $listboxsize = $lb_out->size;
    my $size        = $listboxsize - 1;

    $lb_out->delete( 0, $size );

}

# 'Information'  action
sub getInfo {
    my $info = $mw->DialogBox(
        -title   => "About this program",
        -buttons => ["OK"]
    );

    $info->add(
        'Label',
        -anchor     => 'w',
        -justify    => 'left',
        -background => '#FFFFFF',
        -text       => qq(
Excel Metadata to EML - Version 0.1

The Excel Metadata to EML program converts LTER EML Metadata Submission Template files 
(in Excel format) to EML 2.0.1 files.  

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

This material is based upon work supported by the National Science Foundation 
under Cooperative Agreement #DEB-9910514. Any opinions, findings, conclusions, 
or recommendations expressed in the material are those of the author(s) and
 do not necessarily reflect the views of the National Science Foundation.

Copyright (C) 2004  Florida International University

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
)
    )->pack;

    $info->Show();

}

# 'Instructions'  action
sub getInstructions {
    my $info = $mw->DialogBox(
        -title   => "Instructions",
        -buttons => ["OK"]
    );

    $info->add(
        'Label',
        -anchor     => 'w',
        -justify    => 'left',
        -background => '#FFFFFF',
        -text       => qq(
Excel Metadata to EML - Version 0.1

This program converts LTER Excel Metadata Templates to EML 2.0.1 files.  The metadata template and
this program are based on the EML best practices document released in 2004.  

INSTRUCTIONS

    1. Fill out the LTER EML Metadata Template (xls_eml_01.xls).  
    Instructions for filling out the template are provided in a Microsoft Word help document called xls_eml_01.doc

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
       Please note that EML 2.0.1 or higher is required to validate EML files with embedded data.

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
    error log file (error.log). Files in the list which don't have xls extensions or the value 'Dataset Title' 
    in cell B21 will not be converted to EML.  Click on the 'Stop converting files' button or select 
    'Stop converting files' from the file menu if you wish to stop the conversion process.

)
    )->pack;

    $info->Show();

}

# 'Notes'  action
sub getNotes {
    my $info = $mw->DialogBox(
        -title   => "Notes",
        -buttons => ["OK"]
    );

    $info->add(
        'Label',
        -anchor     => 'w',
        -justify    => 'left',
        -background => '#FFFFFF',
        -text       => qq(
Excel Metadata to EML - Version 0.1

NOTES

    - EML files will retain the same name as the Excel files, but their file extension will be 'xml' instead of 'xls'. 

    - The program will embed data entered in the Values section of the DataTable worksheet into the EML file if the 
    'Embed data' box is checked.  If no values are entered in this section, no data will be embedded in the EML. 
    If you plan to embed data, please use EML 2.0.1 or higher.

    - Some or all buttons may not be visible if your screen resolution is less than 1024 X 768.  If buttons aren't
    visible, try making the window larger or using the commands in the File menu at the top instead.

    - The program uses up a chunk of memory each time it converts an Excel file to EML.  This memory isn't returned to
    the system until after you exit the program.

       Tips for best performance:

       - Validate against local schema or deselect the validation option.  Validation against a schema URL can take 
       awhile if you have a slow network connection (this is very slow with a dial-up connection!).

       - Keep the size of the Excel Metadata Template files as small as possible.  For example, if you don't have 
       any methods citations, you could clear all of the cells in the methodsCitation sheet to reduce the size of the file.
       Please note that every worksheet needs to be present in the specified order for the program to work, though.

)
    )->pack;

    $info->Show();

}

#######################################################################################################
# 'Convert all files to EML' button action                                                            #
#  -Checks the left listbox for file paths - make sure the file list contains files                   #
#  -Checks file paths for xls extension - make sure they're Excel files                               #
#                                                                                                     #
#  -Checks Excel files for 'Dataset Title' value in cell B21 - make sure they're metadata templates   #
# If the file passes all of the tests, the file path, schema, indent level, and destination directory #
# are passed to the createEMLFile subroutine. Otherwise, a message appears in the log explaining      #
# why the file wasn't processed.                                                                      #
#######################################################################################################

my $total_files;
my $files_done;
my $EML_file_done;

sub convertToEML {
    my @file_list   = ();
    my $listboxsize = $lb_in->size;
    my $size        = $listboxsize - 1;
    my $end_loop;
    $loop = 1;

    if ( !$listboxsize ) {
        $lb_out->insert( "end", ":-O  No files have been added to the list yet." );
        $lb_out->insert( "end", " " );
    }
    else {

        @file_list = $lb_in->get( 0, 'end' );

        my $item;
        $percent_done  = 0;
        $files_done    = 0;
        $EML_file_done = 0;
        $progress->configure(-value => $percent_done);
        $progress->update;

        $total_files = $#file_list + 1;

        foreach $item (@file_list) {
            if ($loop) {
                chomp $item;
                my $filename = $item;
                my $eml_filename;
                my $xls_filename;
                my $filetest;
                if ( $filename =~ /\// ) {
                    my @xls_file_path = split( /\//, $filename );
                    $xls_filename = $xls_file_path[$#xls_file_path];
                    my $eml_file = $xls_filename;
                    chop($eml_file);
                    chop($eml_file);
                    chop($eml_file);
                    chop($eml_file);
                    $eml_filename = $eml_file . "\.xml";
                }
                else {
                    my @xls_file_path = split( /\\/, $filename );
                    $xls_filename = $xls_file_path[$#xls_file_path];
                    my $eml_file = $xls_filename;
                    chop($eml_file);
                    chop($eml_file);
                    chop($eml_file);
                    chop($eml_file);
                    $eml_filename = $eml_file . "\.xml";
                }

                my @eml_file = split( /\./, $filename );

                # Check for xls extension

                if ( $eml_file[$#eml_file] eq 'xls' ) {

                    # Check for value in cell B21 (should be Dataset Title)

                    my $Excel = new Spreadsheet::ParseExcel;
                    my $Book  = Spreadsheet::ParseExcel::Workbook->Parse($filename);
                    my $WkS0  = $Book->{Worksheet}[0];
                    my $WkS1  = $Book->{Worksheet}[1];

                    if ( $WkS0->{Cells}[20][1] ) {
                        $filetest = $WkS0->{Cells}[20][1]->Value;

                        if ( $filetest eq 'Dataset Title' ) {

                            if ( !$save_dir ) {
                                $save_dir = "";
                            }
                            if ( !$schema ) {
                                $schema = "http://cvs.lternet.edu/cgi-bin/viewcvs.cgi/*checkout*/eml/eml-2.0.1/eml.xsd";
                            }
                            if ( !$stmml ) {
                                $stmml = "http://cvs.lternet.edu/cgi-bin/viewcvs.cgi/*checkout*/eml/eml-2.0.1/stmml.xsd";
                            }
                            if ( !$stylesheet ) {
                                $stylesheet = "";
                            }
                            if ( !$indent_level ) {
                                $indent_level = "2";
                            }

                            $lb_out->insert( "end", "Starting to convert $xls_filename to $eml_filename" );
                            my $files_done2 = &createEMLFile( $filename, $save_dir, $schema, $indent_level, $total_files, $files_done, $stmml, $stylesheet, $embed_data_checkbox );
                            $files_done   = $files_done2;
                            $percent_done = ( $files_done / $total_files ) * 100;

                            $progress->configure(-value => $percent_done);
                            $progress->update;
                        }
                        else {
                            $files_done   = $files_done + 1;
                            $percent_done = ( $files_done / $total_files ) * 100;
                            $progress->configure(-value => $percent_done);
                            $progress->update;
                            $lb_out->insert( "end", ":-O  " . "$xls_filename" . " does not seem to be an Excel Metadata file." );
                            $lb_out->insert( "end", "      (cell B21 should be Dataset Title)" );
                            $lb_out->insert( "end", " " );

                        }
                    }
                    else {
                        $files_done   = $files_done + 1;
                        $percent_done = ( $files_done / $total_files ) * 100;
                        $progress->configure(-value => $percent_done);
                        $progress->update;
                        $lb_out->insert( "end", ":-O  " . "$xls_filename" . " does not seem to be an Excel Metadata file." );
                        $lb_out->insert( "end", "      (cell B21 should be Dataset Title)" );
                        $lb_out->insert( "end", " " );
                    }
                }
                else {
                    $files_done   = $files_done + 1;
                    $percent_done = ( $files_done / $total_files ) * 100;
                    $progress->configure(-value => $percent_done);
                    $progress->update;
                    $lb_out->insert( "end", ":-O  " . "$xls_filename" . " does not have an xls extension." );
                    $lb_out->insert( "end", " " );
                }

                $percent_done = ( $files_done / $total_files ) * 100;
                $progress->configure(-value => $percent_done);
                $progress->update;
            }
            else {
                if ($end_loop) {
                }
                else {
                    $percent_done = 100;
                    $lb_out->insert( "end", "EML conversion stopped!" );
                    $lb_out->insert( "end", " " );
                    $progress->configure(-value => $percent_done);
                    $progress->update;
                    $end_loop = 1;
                }
            }

        }

    }
}

######################################
# Subroutine to create the EML file  #
######################################

sub createEMLFile {

    ############################################
    # Determine destination file name and path #
    ############################################

    my $filename = $_[0];
    my $eml_file;
    if ( $filename =~ /\// && $_[1] gt '' ) {
        my @xls_file_path = split( /\//, $filename );
        my $xls_filename = $xls_file_path[$#xls_file_path];
        chop($xls_filename);
        chop($xls_filename);
        chop($xls_filename);
        chop($xls_filename);
        $eml_file = $_[1] . "/" . $xls_filename . "\.xml";
    }
    elsif ( $filename =~ /\\/ && $_[1] gt '' ) {
        my @xls_file_path = split( /\\/, $filename );
        my $xls_filename = $xls_file_path[$#xls_file_path];
        chop($xls_filename);
        chop($xls_filename);
        chop($xls_filename);
        chop($xls_filename);
        $eml_file = $_[1] . "\\" . $xls_filename . "\.xml";
    }
    else {
        $eml_file = $filename;
        chop($eml_file);
        chop($eml_file);
        chop($eml_file);
        chop($eml_file);
        $eml_file = $eml_file . "\.xml";
        $lb_out->insert( "end", "      A directory in which to save the EML file was not specified." );
        $lb_out->insert( "end", "      The EML file will be saved in the same directory as its source Excel file." );
    }

    my $schemalocation      = $_[2];
    my $stmml               = $_[6];
    my $stylesheet          = $_[7];
    my $embed_data_checkbox = $_[8];
    my $indent              = " " x $_[3];
    $main::indent = $indent;

    $total_files  = $_[4];
    $files_done   = $_[5];
    $percent_done = ( $files_done / $total_files ) * 100;
    $progress->configure(-value => $percent_done);
    $progress->update;

    my $Excel = new Spreadsheet::ParseExcel;
    my $Book  = Spreadsheet::ParseExcel::Workbook->Parse($filename);
    my $WkS0  = $Book->{Worksheet}[0];
    my $WkS1  = $Book->{Worksheet}[1];
    my $WkS2  = $Book->{Worksheet}[2];
    my $WkS3  = $Book->{Worksheet}[3];
    my $WkS4  = $Book->{Worksheet}[4];

    # Subroutine to calculate percent_done for progress bar
    sub percentDone {
        $files_done   = $files_done + .1;
        $percent_done = ( $files_done / $total_files ) * 100;
        $progress->configure(-value => $percent_done);
        $progress->update;
        return $files_done;
    }

    # Subroutine to replace &, <, >, µ with entities
    # I'm not sure if this is necessary or a good idea
    sub IllegalChars_Array {
        my $value;
        foreach $value (@_) {
            $value =~ s/&/&amp\;/g;
            $value =~ s/>/&gt\;/g;
            $value =~ s/</&lt\;/g;
            $value =~ s/µ/&#181;/g;
        }
    }

    ######################################################################
    # Subroutines to get strings and arrays from cells in the Excel file #
    ######################################################################

    # Gets the information from only one column (one cell)
    sub getStringValue {

        if ( $_[0] ) {
            my $value;
            $value = $_[0]->Value;

            $value =~ s/&/&amp\;/g;
            $value =~ s/>/&gt\;/g;
            $value =~ s/</&lt\;/g;
            $value =~ s/µ/&#181;/g;

            return $value;
        }

    }

    # Gets information (divided into sections by |) from one column (one cell) and splits it into an array
    sub getArrayValue {

        if ( $_[0] ) {
            my @value;
            my $value;
            @value = split( /\|/, $_[0]->Value );

            &IllegalChars_Array(@value);
            return @value;
        }

    }

    # Gets information from one to multiple columns (information not part of a group)
    sub getArrayValueColumns {
        my $row    = $_[0];
        my $column = $_[1];
        my $WkS    = $_[2];
        my @value;
        my $value;
        my $value_test;

        while ( $WkS->{Cells}[$row][$column] ) {

            $value_test = $WkS->{Cells}[$row][$column];
            if ($value_test) {
                $value = $WkS->{Cells}[$row][$column]->Value;
                if ( $value gt '' ) {
                    push( @value, $value );
                }
            }
            $column = $column + 1;
        }

        &IllegalChars_Array(@value);
        return @value;
    }

    # Gets the maximum number of columns used for a group of rows which might have multiple columns
    sub getNumGroupColumns {

        my $WkS = $_[3];
        my %group_columns;
        my $group_col = $_[2];
        my $group_row;
        my $continue = 1;
        my $columns;

        while ($continue) {
            my $group_rows_start = ( $_[0] );
            my $group_rows_end   = ( $_[1] );
            while ( $group_rows_start <= $group_rows_end ) {
                my $value = $WkS->{Cells}[$group_rows_start][$group_col];
                if ($value) {
                    $value = $WkS->{Cells}[$group_rows_start][$group_col]->Value;
                    if ( $value gt '' ) {
                        if ( exists $group_columns{$group_col} ) {
                        }
                        else {
                            $group_columns{$group_col} = 1;
                        }
                    }
                    else {
                    }
                }
                $group_rows_start = $group_rows_start + 1;
            }
            if ( exists $group_columns{$group_col} ) {
                $group_col = $group_col + 1;
                $continue  = 1;
            }
            else {
                $continue  = 0;
                $group_col = $group_col - 1;
            }

        }
        return $group_col;
    }

    # Gets information from a row which is part of group with multiple columns.
    # Uses the maximum number of columns in a group to make sure that each position
    # in each array in the group refers to the same column (ie. $row1[3], $row2[3],
    # $row3[3] all refer to cells in the same column).

    sub getGroupedColumns {
        my $count          = $_[3];
        my $count_finished = $_[0];
        my $worksheet      = $_[2];
        my @value;
        my $value;

        while ( $count <= $count_finished ) {
            if ( $worksheet->{Cells}[ $_[1] ][$count] ) {
                push( @value, $worksheet->{Cells}[ $_[1] ][$count]->Value );
            }
            else {
                push( @value, "" );
            }
            $count = $count + 1;
        }
        &IllegalChars_Array(@value);

        return @value;

    }

    # Gets data values that will be embedded in EML
    sub getEmbeddedData {
        my $delimiter  = $_[0];
        my $WkS        = $_[1];
        my $WkS1       = $_[2];
        my $row        = 33;
        my $column     = 1;
        my $column_end = "no";
        my $row_end    = "no";
        my @value;
        my $value;
        my $embedded_data;

        if ( $delimiter eq 'comma' || $delimiter eq 'Comma' || $delimiter eq 'COMMA' ) {
            $delimiter = ",";
        }
        elsif ( $delimiter eq 'tab' || $delimiter eq 'Tab' || $delimiter eq 'TAB' ) {
            $delimiter = "\t";
        }
        else {
            $delimiter = "\t";
        }

        while ( ( $WkS->{Cells}[$row][$column] || $row_end eq 'no' ) && $delimiter ) {
            if ( $column_end eq 'no' ) {

                if ( $WkS->{Cells}[$row][$column] ) {
                    $value = $WkS->{Cells}[$row][$column]->Value;

                    if ($value) {

                        if ($embedded_data) {
                            $embedded_data = $embedded_data . $delimiter . $value;
                        }
                        else {
                            $embedded_data = $value;
                        }

                    }
                    $column     = $column + 1;
                    $column_end = "no";

                }
                else {
                    $column_end = "yes";
                }
            }
            elsif ( $column_end eq 'yes' ) {
                $column = 1;
                $row    = $row + 1;
                push( @value, $embedded_data );
                $column_end    = "no";
                $embedded_data = undef;
                if ( $WkS->{Cells}[$row][$column] ) {
                    $value = $WkS->{Cells}[$row][$column]->Value;

                    if ($value) {
                        $row_end = "no";
                    }
                    else {
                        $row_end = "yes";
                    }

                }
            }
            else {
                @value = ();
            }
        }

        &IllegalChars_Array(@value);
        return @value;
    }

    # Prints start tag (with an option to print an ID in the start tag), string, and end tag
    sub printXMLString {
        if ( $_[0] ) {
            my $emltag   = $_[1];
            my $emlvalue = $_[0];
            my $level    = $_[2];
            my $spaces   = $main::indent x $level;
            my $id       = $_[3];
            my $idname   = $_[4];

            print XML "$spaces<$emltag";
            if ($id) {
                if ($idname) {
                    print XML " $idname=\"$id\">";
                }
                else {
                    print XML " id=\"$id\">";
                }
            }
            else {
                print XML ">";
            }
            print XML "$emlvalue";
            print XML "</$emltag>\n";
        }
    }

    # Only prints start tag
    sub printXMLStartTag {

        my $emltag = $_[0];
        my $level  = $_[1];
        my $spaces = $main::indent x $level;
        my $id     = $_[2];
        my $idname = $_[3];

        print XML "$spaces<$emltag";
        if ($id) {
            if ($idname) {
                print XML " $idname=\"$id\">\n";
            }
            else {
                print XML " id=\"$id\">\n";
            }
        }
        else {
            print XML ">\n";
        }
    }

    # Only prints an end tag
    sub printXMLEndTag {

        my $emltag = $_[0];
        my $level  = $_[1];
        my $spaces = $main::indent x $level;

        print XML "$spaces</$emltag>\n";
    }

    #################################################################################################
    # Ten percentDone subroutines are scattered through the createEMLFile for the progress dialog   #
    # Each percentDone subroutine adds 0.1 to $files_done (indicates another 10% of a file is done) #
    #################################################################################################

    ############
    percentDone;
    ############

    #################################################################################
    #  Using subroutines above to store metadata information in strings and arrays  #
    #################################################################################

    my $siteabbrev     = getStringValue( $WkS0->{Cells}[17][2] );
    my $metacat_pkg_id = getStringValue( $WkS0->{Cells}[18][2] );
    my $DatasetID      = getStringValue( $WkS0->{Cells}[19][2] );
    my $dataset_title  = getStringValue( $WkS0->{Cells}[20][2] );

    my $creator_rows_start   = 21;
    my $creator_rows_end     = 34;
    my $creator_column_start = 2;
    my $creator_columns      = getNumGroupColumns( $creator_rows_start, $creator_rows_end, 2, $WkS0 );
    my @creator_salutation   = getGroupedColumns( $creator_columns, 21, $WkS0, 2 );
    my @creator_firstname    = getGroupedColumns( $creator_columns, 22, $WkS0, 2 );
    my @creator_lastname     = getGroupedColumns( $creator_columns, 23, $WkS0, 2 );
    my @creator_organization = getGroupedColumns( $creator_columns, 24, $WkS0, 2 );
    my @creator_position     = getGroupedColumns( $creator_columns, 25, $WkS0, 2 );
    my @creator_address      = getGroupedColumns( $creator_columns, 26, $WkS0, 2 );
    my @creator_city         = getGroupedColumns( $creator_columns, 27, $WkS0, 2 );
    my @creator_state        = getGroupedColumns( $creator_columns, 28, $WkS0, 2 );
    my @creator_zipcode      = getGroupedColumns( $creator_columns, 29, $WkS0, 2 );
    my @creator_country      = getGroupedColumns( $creator_columns, 30, $WkS0, 2 );
    my @creator_phone        = getGroupedColumns( $creator_columns, 31, $WkS0, 2 );
    my @creator_fax          = getGroupedColumns( $creator_columns, 32, $WkS0, 2 );
    my @creator_email        = getGroupedColumns( $creator_columns, 33, $WkS0, 2 );
    my @creator_url          = getGroupedColumns( $creator_columns, 34, $WkS0, 2 );

    ############
    percentDone;
    ############

    my @dataset_abstract          = getArrayValueColumns( 36, 2, $WkS0 );
    my @dataset_keywords          = getArrayValueColumns( 37, 2, $WkS0 );
    my @dataset_keyword_thesaurus = getArrayValueColumns( 38, 2, $WkS0 );

    my $geodesc_rows_start        = 39;
    my $geodesc_rows_end          = 43;
    my $geodesc_column_start      = 2;
    my $geodesc_columns           = getNumGroupColumns( $geodesc_rows_start, $geodesc_rows_end, 2, $WkS0 );
    my @geographic_description    = getGroupedColumns( $geodesc_columns, 39, $WkS0, 2 );
    my @data_west_bounding_coord  = getGroupedColumns( $geodesc_columns, 40, $WkS0, 2 );
    my @data_east_bounding_coord  = getGroupedColumns( $geodesc_columns, 41, $WkS0, 2 );
    my @data_north_bounding_coord = getGroupedColumns( $geodesc_columns, 42, $WkS0, 2 );
    my @data_south_bounding_coord = getGroupedColumns( $geodesc_columns, 43, $WkS0, 2 );

    my $data_entity_beginning_temporal_coverage_date = getStringValue( $WkS0->{Cells}[44][2] );
    my $data_entity_ending_temporal_coverage_date    = getStringValue( $WkS0->{Cells}[45][2] );

    my $dataent_taxon_rows_start       = 46;
    my $dataent_taxon_rows_end         = 48;
    my $dataent_taxon_column_start     = 2;
    my $dataent_taxon_columns          = getNumGroupColumns( $dataent_taxon_rows_start, $dataent_taxon_rows_end, 2, $WkS0 );
    my @data_entity_taxon_rank_name    = getGroupedColumns( $dataent_taxon_columns, 46, $WkS0, 2 );
    my @data_entity_taxon_rank_value   = getGroupedColumns( $dataent_taxon_columns, 47, $WkS0, 2 );
    my @data_entity_common_taxon_names = getGroupedColumns( $dataent_taxon_columns, 48, $WkS0, 2 );

    my @dataset_intellectual_rights = getArrayValueColumns( 50, 2, $WkS0 );
    my $dataset_download_url        = getStringValue( $WkS0->{Cells}[51][2] );

    my $dataset_offline_medium_name          = getStringValue( $WkS0->{Cells}[52][2] );
    my $dataset_offline_medium_density       = getStringValue( $WkS0->{Cells}[53][2] );
    my $dataset_offline_medium_density_units = getStringValue( $WkS0->{Cells}[54][2] );
    my $dataset_offline_medium_volume        = getStringValue( $WkS0->{Cells}[55][2] );
    my $dataset_offline_medium_format        = getStringValue( $WkS0->{Cells}[56][2] );

    my $assocparty_rows_start   = 58;
    my $assocparty_rows_end     = 70;
    my $assocparty_column_start = 2;
    my $assocparty_columns      = getNumGroupColumns( $assocparty_rows_start, $assocparty_rows_end, 2, $WkS0 );
    my @assocparty_firstname    = getGroupedColumns( $assocparty_columns, 58, $WkS0, 2 );
    my @assocparty_lastname     = getGroupedColumns( $assocparty_columns, 59, $WkS0, 2 );
    my @assocparty_organization = getGroupedColumns( $assocparty_columns, 60, $WkS0, 2 );
    my @assocparty_address      = getGroupedColumns( $assocparty_columns, 61, $WkS0, 2 );
    my @assocparty_city         = getGroupedColumns( $assocparty_columns, 62, $WkS0, 2 );
    my @assocparty_state        = getGroupedColumns( $assocparty_columns, 63, $WkS0, 2 );
    my @assocparty_zipcode      = getGroupedColumns( $assocparty_columns, 64, $WkS0, 2 );
    my @assocparty_country      = getGroupedColumns( $assocparty_columns, 65, $WkS0, 2 );
    my @assocparty_phone        = getGroupedColumns( $assocparty_columns, 66, $WkS0, 2 );
    my @assocparty_fax          = getGroupedColumns( $assocparty_columns, 67, $WkS0, 2 );
    my @assocparty_email        = getGroupedColumns( $assocparty_columns, 68, $WkS0, 2 );
    my @assocparty_role         = getGroupedColumns( $assocparty_columns, 69, $WkS0, 2 );
    my @assocparty_url          = getGroupedColumns( $assocparty_columns, 70, $WkS0, 2 );

    my $contact_rows_start   = 72;
    my $contact_rows_end     = 84;
    my $contact_column_start = 2;
    my $contact_columns      = getNumGroupColumns( $contact_rows_start, $contact_rows_end, 2, $WkS0 );
    my @contact_firstname    = getGroupedColumns( $contact_columns, 72, $WkS0, 2 );
    my @contact_lastname     = getGroupedColumns( $contact_columns, 73, $WkS0, 2 );
    my @contact_organization = getGroupedColumns( $contact_columns, 74, $WkS0, 2 );
    my @contact_position     = getGroupedColumns( $contact_columns, 75, $WkS0, 2 );
    my @contact_address      = getGroupedColumns( $contact_columns, 76, $WkS0, 2 );
    my @contact_city         = getGroupedColumns( $contact_columns, 77, $WkS0, 2 );
    my @contact_state        = getGroupedColumns( $contact_columns, 78, $WkS0, 2 );
    my @contact_zipcode      = getGroupedColumns( $contact_columns, 79, $WkS0, 2 );
    my @contact_country      = getGroupedColumns( $contact_columns, 80, $WkS0, 2 );
    my @contact_phone        = getGroupedColumns( $contact_columns, 81, $WkS0, 2 );
    my @contact_fax          = getGroupedColumns( $contact_columns, 82, $WkS0, 2 );
    my @contact_email        = getGroupedColumns( $contact_columns, 83, $WkS0, 2 );
    my @contact_url          = getGroupedColumns( $contact_columns, 84, $WkS0, 2 );

    my $publisher_rows_start   = 86;
    my $publisher_rows_end     = 94;
    my $publisher_column_start = 2;
    my $publisher_columns      = getNumGroupColumns( $publisher_rows_start, $publisher_rows_end, 2, $WkS0 );
    my @publisher_organization = getGroupedColumns( $publisher_columns, 86, $WkS0, 2 );
    my @publisher_address      = getGroupedColumns( $publisher_columns, 87, $WkS0, 2 );
    my @publisher_city         = getGroupedColumns( $publisher_columns, 88, $WkS0, 2 );
    my @publisher_state        = getGroupedColumns( $publisher_columns, 89, $WkS0, 2 );
    my @publisher_zipcode      = getGroupedColumns( $publisher_columns, 90, $WkS0, 2 );
    my @publisher_country      = getGroupedColumns( $publisher_columns, 91, $WkS0, 2 );
    my @publisher_phone        = getGroupedColumns( $publisher_columns, 92, $WkS0, 2 );
    my @publisher_email        = getGroupedColumns( $publisher_columns, 93, $WkS0, 2 );
    my @publisher_url          = getGroupedColumns( $publisher_columns, 94, $WkS0, 2 );

    my $mdprovider_rows_start   = 96;
    my $mdprovider_rows_end     = 104;
    my $mdprovider_column_start = 2;
    my $mdprovider_columns      = getNumGroupColumns( $mdprovider_rows_start, $mdprovider_rows_end, 2, $WkS0 );
    my @mdprovider_organization = getGroupedColumns( $mdprovider_columns, 96, $WkS0, 2 );
    my @mdprovider_address      = getGroupedColumns( $mdprovider_columns, 97, $WkS0, 2 );
    my @mdprovider_city         = getGroupedColumns( $mdprovider_columns, 98, $WkS0, 2 );
    my @mdprovider_state        = getGroupedColumns( $mdprovider_columns, 99, $WkS0, 2 );
    my @mdprovider_zipcode      = getGroupedColumns( $mdprovider_columns, 100, $WkS0, 2 );
    my @mdprovider_country      = getGroupedColumns( $mdprovider_columns, 101, $WkS0, 2 );
    my @mdprovider_phone        = getGroupedColumns( $mdprovider_columns, 102, $WkS0, 2 );
    my @mdprovider_email        = getGroupedColumns( $mdprovider_columns, 103, $WkS0, 2 );
    my @mdprovider_url          = getGroupedColumns( $mdprovider_columns, 104, $WkS0, 2 );

    my $dataset_publication_date           = getStringValue( $WkS0->{Cells}[106][2] );
    my $dataset_access_authentication_info = getStringValue( $WkS0->{Cells}[107][2] );
    my @dataset_principal_access_info      = getArrayValueColumns( 108, 2, $WkS0 );
    my @dataset_principal_permission_info  = getArrayValueColumns( 109, 2, $WkS0 );

    my $dataset_methods_rows_start   = 111;
    my $dataset_methods_rows_end     = 114;
    my $dataset_methods_column_start = 2;
    my $dataset_methods_columns      = getNumGroupColumns( $dataset_methods_rows_start, $dataset_methods_rows_end, 2, $WkS0 );
    my @dataset_methods_desc         = getGroupedColumns( $dataset_methods_columns, 111, $WkS0, 2 );
    my @dataset_methods_citationID   = getGroupedColumns( $dataset_methods_columns, 112, $WkS0, 2 );
    my @dataset_methods_protocolID   = getGroupedColumns( $dataset_methods_columns, 113, $WkS0, 2 );
    my @dataset_methods_instrument   = getGroupedColumns( $dataset_methods_columns, 114, $WkS0, 2 );

    my @dataset_sampling_desc = getArrayValue( $WkS0->{Cells}[116][2] );
    my @dataset_studyext_desc = getArrayValue( $WkS0->{Cells}[117][2] );

    my $sampling_sites_geodesc_rows_start     = 119;
    my $sampling_sites_geodesc_rows_end       = 125;
    my $sampling_sites_geodesc_column_start   = 2;
    my $sampling_sites_geodesc_columns        = getNumGroupColumns( $sampling_sites_geodesc_rows_start, $sampling_sites_geodesc_rows_end, 2, $WkS0 );
    my @sampling_sites_geographic_description = getGroupedColumns( $sampling_sites_geodesc_columns, 119, $WkS0, 2 );
    my @sampling_sites_west_bounding_coord    = getGroupedColumns( $sampling_sites_geodesc_columns, 120, $WkS0, 2 );
    my @sampling_sites_east_bounding_coord    = getGroupedColumns( $sampling_sites_geodesc_columns, 121, $WkS0, 2 );
    my @sampling_sites_north_bounding_coord   = getGroupedColumns( $sampling_sites_geodesc_columns, 122, $WkS0, 2 );
    my @sampling_sites_south_bounding_coord   = getGroupedColumns( $sampling_sites_geodesc_columns, 123, $WkS0, 2 );
    my @sampling_sites_latitude               = getGroupedColumns( $sampling_sites_geodesc_columns, 124, $WkS0, 2 );
    my @sampling_sites_longitude              = getGroupedColumns( $sampling_sites_geodesc_columns, 125, $WkS0, 2 );

    my @dataset_quality_control_info = getArrayValue( $WkS0->{Cells}[127][2] );
    my @dataset_maintenance          = getArrayValue( $WkS0->{Cells}[128][2] );

    ############
    percentDone;
    ############

    my @data_entity_name           = getArrayValue( $WkS0->{Cells}[133][2] );
    my @data_entity_desc           = getArrayValue( $WkS0->{Cells}[134][2] );
    my $data_object_name           = getStringValue( $WkS0->{Cells}[135][2] );
    my $num_data_records           = getStringValue( $WkS0->{Cells}[136][2] );
    my $num_header_lines           = getStringValue( $WkS0->{Cells}[137][2] );
    my $data_attribute_orientation = getStringValue( $WkS0->{Cells}[138][2] );
    my $data_field_delimiter       = getStringValue( $WkS0->{Cells}[139][2] );
    my $data_external_format       = getStringValue( $WkS0->{Cells}[140][2] );

    my @research_project_number = getArrayValue( $WkS0->{Cells}[145][2] );

    my $attribute_rows_start      = 10;
    my $attribute_rows_end        = 32;
    my $attribute_column_start    = 1;
    my $attribute_columns         = getNumGroupColumns( $attribute_rows_start, $attribute_rows_end, 1, $WkS4 );
    my @attribute_name            = getGroupedColumns( $attribute_columns, 10, $WkS4, 1 );
    my @attribute_label           = getGroupedColumns( $attribute_columns, 11, $WkS4, 1 );
    my @attribute_definition      = getGroupedColumns( $attribute_columns, 12, $WkS4, 1 );
    my @missing_value_code        = getGroupedColumns( $attribute_columns, 13, $WkS4, 1 );
    my @missing_value_explanation = getGroupedColumns( $attribute_columns, 14, $WkS4, 1 );
    my @measurement_scale         = getGroupedColumns( $attribute_columns, 15, $WkS4, 1 );
    my @codeset_name              = getGroupedColumns( $attribute_columns, 16, $WkS4, 1 );
    my @number_type               = getGroupedColumns( $attribute_columns, 17, $WkS4, 1 );
    my @variable_type             = getGroupedColumns( $attribute_columns, 18, $WkS4, 1 );
    my @date_time_format          = getGroupedColumns( $attribute_columns, 19, $WkS4, 1 );
    my @date_time_min             = getGroupedColumns( $attribute_columns, 20, $WkS4, 1 );
    my @date_time_max             = getGroupedColumns( $attribute_columns, 21, $WkS4, 1 );
    my @units_data_table          = getGroupedColumns( $attribute_columns, 22, $WkS4, 1 );
    my @units                     = getGroupedColumns( $attribute_columns, 23, $WkS4, 1 );
    my @custom_or_eml             = getGroupedColumns( $attribute_columns, 24, $WkS4, 1 );
    my @custom_unitType           = getGroupedColumns( $attribute_columns, 25, $WkS4, 1 );
    my @custom_unitID             = getGroupedColumns( $attribute_columns, 26, $WkS4, 1 );
    my @custom_unitParentSI       = getGroupedColumns( $attribute_columns, 27, $WkS4, 1 );
    my @custom_unitMultiplierToSI = getGroupedColumns( $attribute_columns, 28, $WkS4, 1 );
    my @custom_unitAbrev          = getGroupedColumns( $attribute_columns, 29, $WkS4, 1 );
    my @custom_unitDesc           = getGroupedColumns( $attribute_columns, 30, $WkS4, 1 );
    my @precision                 = getGroupedColumns( $attribute_columns, 31, $WkS4, 1 );
    my @calculations              = getGroupedColumns( $attribute_columns, 32, $WkS4, 1 );
    my @custom_unit_list;
    my @custom_unit_stmml_tag;
    my @custom_unit_stmml_desc_tag;

    my @embedded_data;
    my $embedded_data_test = getStringValue( $WkS4->{Cells}[33][1] );

    if ( $embedded_data_test && ( $embed_data_checkbox eq 'yes' ) ) {
        @embedded_data = getEmbeddedData( $data_field_delimiter, $WkS4, $WkS1 );
    }
    else {
    }

    sub directory_die {

        $lb_out->insert( "end", "  " );
        $lb_out->insert( "end", ":-O  Can't write to the specified destination directory." );
        $lb_out->insert( "end", "       ($_[0])." );
        $lb_out->insert( "end", "       Please verify that this directory exists and that you can write to the directory." );
        $percent_done = 100;
        $progress->configure(-value => $percent_done);
        $progress->update;

    }

    open( XML, ">$eml_file" ) or directory_die($eml_file);

    ############
    percentDone;
    ############

    ###############################
    # Start printing the XML file #
    ###############################

    print XML "<?xml version=\"1.0\"  encoding=\"UTF\-8\"?>\n";
    if ($stylesheet) {
        print XML "<?xml-stylesheet href=\"$stylesheet\" type=\"text/xsl\"?>\n";
    }
    print XML
"<eml:eml packageId=\"knb-lter-$metacat_pkg_id\" system=\"knb\" xsi:schemaLocation=\"eml://ecoinformatics.org/eml-2.0.1 $schemalocation\"\n xmlns:ds=\"eml://ecoinformatics.org/dataset-2.0.1\"\n xmlns:eml=\"eml://ecoinformatics.org/eml-2.0.1\" \n xmlns:stmml=\"http://www.xml-cml.org/schema/stmml\" \n xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n";

    printXMLStartTag( "dataset", "1", $DatasetID );
    printXMLString( $DatasetID, "alternateIdentifier", "2" );
    printXMLString( $dataset_title, "title", "2" );

    #################################
    # Print DATASET CREATOR SECTION #
    #################################

    my $creator = 0;

    while ( $creator <= ( $creator_columns - $creator_column_start ) ) {

        printXMLStartTag( "creator", "2" );

        if ( $creator_salutation[$creator] || $creator_firstname[$creator] || $creator_lastname[$creator] ) {
            printXMLStartTag( "individualName", "3" );
            printXMLString( $creator_salutation[$creator], "salutation", "4" );
            printXMLString( $creator_firstname[$creator],  "givenName",  "4" );
            printXMLString( $creator_lastname[$creator],   "surName",    "4" );
            printXMLEndTag( "individualName", "3" );
        }

        if ( $creator_organization[$creator] ) {
            printXMLString( $creator_organization[$creator], "organizationName", "3" );
        }

        if ( $creator_position[$creator] ) {
            printXMLString( $creator_position[$creator], "positionName", "3" );
        }

        if ( $creator_address[$creator] || $creator_city[$creator] || $creator_state[$creator] || $creator_zipcode[$creator] || $creator_country[$creator] ) {
            printXMLStartTag( "address", "3" );
            my $address = 0;
            my $creator_address_lines;

            my @creator_address_lines = split( /\|/, $creator_address[$creator] );
            while ( $address <= $#creator_address_lines ) {

                printXMLString( $creator_address_lines[$address], "deliveryPoint", "4" );
                $address = $address + 1;
            }
            printXMLString( $creator_city[$creator],    "city",               "4" );
            printXMLString( $creator_state[$creator],   "administrativeArea", "4" );
            printXMLString( $creator_zipcode[$creator], "postalCode",         "4" );
            printXMLString( $creator_country[$creator], "country",            "4" );
            printXMLEndTag( "address", "3" );
        }

        printXMLString( $creator_phone[$creator], "phone", "3", "voice", "phonetype" );
        printXMLString( $creator_fax[$creator],   "phone", "3", "fax",   "phonetype" );
        printXMLString( $creator_email[$creator], "electronicMailAddress", "3" );
        printXMLString( $creator_url[$creator],   "onlineUrl",             "3" );

        printXMLEndTag( "creator", "2" );

        $creator = $creator + 1;
    }

    ###########################################
    # Print DATASET METADATA PROVIDER SECTION #
    ###########################################

    my $mdprovider = 0;

    while ( $mdprovider <= ( $mdprovider_columns - $mdprovider_column_start ) ) {

        printXMLStartTag( "metadataProvider", "2" );

        printXMLString( $mdprovider_organization[$mdprovider], "organizationName", "3" );

        if (   $mdprovider_address[$mdprovider]
            || $mdprovider_city[$mdprovider]
            || $mdprovider_state[$mdprovider]
            || $mdprovider_zipcode[$mdprovider]
            || $mdprovider_country[$mdprovider] )
        {
            printXMLStartTag( "address", "3" );
            my $address = 0;
            my $mdprovider_address_lines;

            my @mdprovider_address_lines = split( /\|/, $mdprovider_address[$mdprovider] );
            while ( $address <= $#mdprovider_address_lines ) {

                printXMLString( $mdprovider_address_lines[$address], "deliveryPoint", "4" );
                $address = $address + 1;
            }
            printXMLString( $mdprovider_city[$mdprovider],    "city",               "4" );
            printXMLString( $mdprovider_state[$mdprovider],   "administrativeArea", "4" );
            printXMLString( $mdprovider_zipcode[$mdprovider], "postalCode",         "4" );
            printXMLString( $mdprovider_country[$mdprovider], "country",            "4" );
            printXMLEndTag( "address", "3" );
        }

        printXMLString( $mdprovider_phone[$mdprovider], "phone", "3", "voice", "phonetype" );
        printXMLString( $mdprovider_email[$mdprovider], "electronicMailAddress", "3" );
        printXMLString( $mdprovider_url[$mdprovider],   "onlineUrl",             "3" );

        printXMLEndTag( "metadataProvider", "2" );

        $mdprovider = $mdprovider + 1;
    }

    ##########################################
    # Print DATASET ASSOCIATED PARTY SECTION #
    ##########################################

    my $assocparty = 0;

    while ( $assocparty <= ( $assocparty_columns - $assocparty_column_start ) ) {

        printXMLStartTag( "associatedParty", "2" );

        printXMLStartTag( "individualName", "3" );
        printXMLString( $assocparty_firstname[$assocparty], "givenName", "4" );
        printXMLString( $assocparty_lastname[$assocparty],  "surName",   "4" );
        printXMLEndTag( "individualName", "3" );
        printXMLString( $assocparty_organization[$assocparty], "organizationName", "3" );

        if (   $assocparty_address[$assocparty]
            || $assocparty_city[$assocparty]
            || $assocparty_state[$assocparty]
            || $assocparty_zipcode[$assocparty]
            || $assocparty_country[$assocparty] )
        {
            printXMLStartTag( "address", "3" );
            my $address = 0;
            my $assocparty_address_lines;

            my @assocparty_address_lines = split( /\|/, $assocparty_address[$assocparty] );
            while ( $address <= $#assocparty_address_lines ) {

                printXMLString( $assocparty_address_lines[$address], "deliveryPoint", "4" );
                $address = $address + 1;
            }
            printXMLString( $assocparty_city[$assocparty],    "city",               "4" );
            printXMLString( $assocparty_state[$assocparty],   "administrativeArea", "4" );
            printXMLString( $assocparty_zipcode[$assocparty], "postalCode",         "4" );
            printXMLString( $assocparty_country[$assocparty], "country",            "4" );
            printXMLEndTag( "address", "3" );
        }
        printXMLString( $assocparty_phone[$assocparty], "phone", "3", "voice", "phonetype" );
        printXMLString( $assocparty_fax[$assocparty],   "phone", "3", "fax",   "phonetype" );
        printXMLString( $assocparty_email[$assocparty], "electronicMailAddress", "3" );
        printXMLString( $assocparty_url[$assocparty],   "onlineUrl",             "3" );
        printXMLString( $assocparty_role[$assocparty],  "role",                  "3" );

        printXMLEndTag( "associatedParty", "2" );

        $assocparty = $assocparty + 1;
    }

    ##################################
    # Print DATASET ABSTRACT SECTION #
    ##################################

    printXMLString( $dataset_publication_date, "pubDate", "2" );

    if (@dataset_abstract) {
        printXMLStartTag( "abstract", "2" );
        my $abstract = 0;
        while ( $abstract <= $#dataset_abstract ) {
            printXMLString( $dataset_abstract[$abstract], "para", "3" );
            $abstract = $abstract + 1;
        }
        printXMLEndTag( "abstract", "2" );
    }

    ##################################
    # Print DATASET KEYWORDS SECTION #
    ##################################

    if (@dataset_keywords) {
        printXMLStartTag( "keywordSet", "2" );
        my $keyword = 0;
        while ( $keyword <= $#dataset_keywords ) {

            printXMLString( $dataset_keywords[$keyword], "keyword", "3" );

            $keyword = $keyword + 1;
        }
        $keyword = 0;
        while ( $keyword <= $#dataset_keywords ) {

            printXMLString( $dataset_keyword_thesaurus[$keyword], "keywordThesaurus", "3" );

            $keyword = $keyword + 1;
        }
        printXMLEndTag( "keywordSet", "2" );
    }

    #############################################
    # Print DATASET INTELLECTUAL RIGHTS SECTION #
    #############################################

    if (@dataset_intellectual_rights) {
        printXMLStartTag( "intellectualRights", "2" );
        my $intellectualRights = 0;

        while ( $intellectualRights <= $#dataset_intellectual_rights ) {
            printXMLString( $dataset_intellectual_rights[$intellectualRights], "para", "3" );
            $intellectualRights = $intellectualRights + 1;
        }
        printXMLEndTag( "intellectualRights", "2" );
    }

    ######################################
    # Print DATASET DISTRIBUTION SECTION #
    ######################################

    if (   $dataset_download_url
        || $dataset_offline_medium_name
        || $dataset_offline_medium_density
        || $dataset_offline_medium_density_units
        || $dataset_offline_medium_volume
        || $dataset_offline_medium_format
        || @embedded_data )
    {
        printXMLStartTag( "distribution", "2" );

        if (@embedded_data) {
            printXMLStartTag( "inline", "3" );

            my $embedded_data_rows = 0;
            while ( $embedded_data_rows <= $#embedded_data ) {

                print XML "$embedded_data[$embedded_data_rows]\n";

                $embedded_data_rows = $embedded_data_rows + 1;
            }

            printXMLEndTag( "inline", "3" );

        }

        if ($dataset_download_url) {
            printXMLStartTag( "online", "3" );
            printXMLString( $dataset_download_url, "url", "4" );
            printXMLEndTag( "online", "3" );
        }

        if (
               $dataset_offline_medium_name
            || $dataset_offline_medium_density
            || $dataset_offline_medium_density_units
            || $dataset_offline_medium_volume
            || $dataset_offline_medium_format

          )
        {
            printXMLStartTag( "offline", "3" );

            printXMLString( $dataset_offline_medium_name,          "mediumName",         "4" );
            printXMLString( $dataset_offline_medium_density,       "mediumDensity",      "4" );
            printXMLString( $dataset_offline_medium_density_units, "mediumDensityUnits", "4" );
            printXMLString( $dataset_offline_medium_volume,        "mediumVolume",       "4" );
            printXMLString( $dataset_offline_medium_format,        "mediumFormat",       "4" );

            printXMLEndTag( "offline", "3" );
        }
        else {
        }

        printXMLEndTag( "distribution", "2" );

    }

    ############
    percentDone;
    ############

    ##########################
    # Print COVERAGE SECTION #
    ##########################

    if (   @geographic_description
        || $data_entity_beginning_temporal_coverage_date
        || $data_entity_ending_temporal_coverage_date
        || @data_entity_taxon_rank_name
        || @data_entity_taxon_rank_value
        || @data_entity_common_taxon_names )
    {

        printXMLStartTag( "coverage", "2" );

        my $geo_coverage = 0;

        if (@geographic_description) {
            while ( $geo_coverage <= ( $geodesc_columns - $geodesc_column_start ) ) {

                printXMLStartTag( "geographicCoverage", "3" );

                printXMLString( $geographic_description[$geo_coverage], "geographicDescription", "4" );

                if (   $data_west_bounding_coord[$geo_coverage]
                    || $data_east_bounding_coord[$geo_coverage]
                    || $data_north_bounding_coord[$geo_coverage]
                    || $data_south_bounding_coord[$geo_coverage] )
                {
                    printXMLStartTag( "boundingCoordinates", "4" );
                    printXMLString( $data_west_bounding_coord[$geo_coverage],  "westBoundingCoordinate",  "5" );
                    printXMLString( $data_east_bounding_coord[$geo_coverage],  "eastBoundingCoordinate",  "5" );
                    printXMLString( $data_north_bounding_coord[$geo_coverage], "northBoundingCoordinate", "5" );
                    printXMLString( $data_south_bounding_coord[$geo_coverage], "southBoundingCoordinate", "5" );
                    printXMLEndTag( "boundingCoordinates", "4" );
                }

                printXMLEndTag( "geographicCoverage", "3" );

                $geo_coverage = $geo_coverage + 1;
            }
        }

        if (   $data_entity_beginning_temporal_coverage_date
            || $data_entity_ending_temporal_coverage_date )
        {
            printXMLStartTag( "temporalCoverage", "3" );
            printXMLStartTag( "rangeOfDates",     "4" );

            if ($data_entity_beginning_temporal_coverage_date) {
                printXMLStartTag( "beginDate", "5" );
                printXMLString( $data_entity_beginning_temporal_coverage_date, "calendarDate", "6" );
                printXMLEndTag( "beginDate", "5" );
            }

            if ($data_entity_ending_temporal_coverage_date) {
                printXMLStartTag( "endDate", "5" );
                printXMLString( $data_entity_ending_temporal_coverage_date, "calendarDate", "6" );
                printXMLEndTag( "endDate", "5" );
            }

            printXMLEndTag( "rangeOfDates",     "4" );
            printXMLEndTag( "temporalCoverage", "3" );
        }

        if (   @data_entity_taxon_rank_name
            || @data_entity_taxon_rank_value
            || @data_entity_common_taxon_names )
        {
            printXMLStartTag( "taxonomicCoverage", "3" );

            my $entity_taxon = 0;
            while ($entity_taxon <= $#data_entity_taxon_rank_name
                || $entity_taxon <= $#data_entity_taxon_rank_value
                || $entity_taxon <= $#data_entity_common_taxon_names )
            {
                printXMLStartTag( "taxonomicClassification", "4" );
                printXMLString( $data_entity_taxon_rank_name[$entity_taxon],    "taxonRankName",  "5" );
                printXMLString( $data_entity_taxon_rank_value[$entity_taxon],   "taxonRankValue", "5" );
                printXMLString( $data_entity_common_taxon_names[$entity_taxon], "taxonCommon",    "5" );
                printXMLEndTag( "taxonomicClassification", "4" );
                $entity_taxon = $entity_taxon + 1;
            }

            printXMLEndTag( "taxonomicCoverage", "3" );
        }

        printXMLEndTag( "coverage", "2" );
    }

    #####################################
    # Print DATASET MAINTENANCE SECTION #
    #####################################

    if (@dataset_maintenance) {
        printXMLStartTag( "maintenance", "2" );

        my $maintenance = 0;

        while ( $maintenance <= $#dataset_maintenance ) {
            printXMLStartTag( "description", "3" );
            printXMLString( $dataset_maintenance[$maintenance], "para", "4" );
            printXMLEndTag( "description", "3" );
            $maintenance = $maintenance + 1;
        }

        printXMLEndTag( "maintenance", "2" );
    }

    #################################
    # Print DATASET CONTACT SECTION #
    #################################

    my $contact = 0;
    while ( $contact <= ( $contact_columns - $contact_column_start ) ) {

        printXMLStartTag( "contact", "2" );

        if ( $contact_firstname[$contact] || $contact_lastname[$contact] ) {
            printXMLStartTag( "individualName", "3" );
            printXMLString( $contact_firstname[$contact], "givenName", "4" );
            printXMLString( $contact_lastname[$contact],  "surName",   "4" );
            printXMLEndTag( "individualName", "3" );
        }

        if ( $contact_organization[$contact] ) {
            printXMLString( $contact_organization[$contact], "organizationName", "3" );
        }

        if ( $contact_position[$contact] ) {
            printXMLString( $contact_position[$contact], "positionName", "3" );
        }

        if ( $contact_address[$contact] || $contact_city[$contact] || $contact_state[$contact] || $contact_zipcode[$contact] || $contact_country[$contact] ) {
            printXMLStartTag( "address", "3" );
            my $address = 0;
            my $contact_address_lines;

            my @contact_address_lines = split( /\|/, $contact_address[$contact] );
            while ( $address <= $#contact_address_lines ) {

                printXMLString( $contact_address_lines[$address], "deliveryPoint", "4" );
                $address = $address + 1;
            }
            printXMLString( $contact_city[$contact],    "city",               "4" );
            printXMLString( $contact_state[$contact],   "administrativeArea", "4" );
            printXMLString( $contact_zipcode[$contact], "postalCode",         "4" );
            printXMLString( $contact_country[$contact], "country",            "4" );
            printXMLEndTag( "address", "3" );
        }
        printXMLString( $contact_phone[$contact], "phone", "3", "voice", "phonetype" );
        printXMLString( $contact_fax[$contact],   "phone", "3", "fax",   "phonetype" );
        printXMLString( $contact_email[$contact], "electronicMailAddress", "3" );
        printXMLString( $contact_url[$contact],   "onlineUrl",             "3" );

        printXMLEndTag( "contact", "2" );

        $contact = $contact + 1;
    }

    ############
    percentDone;
    ############

    ###################################
    # Print DATASET PUBLISHER SECTION #
    ###################################

    my $publisher = 0;

    while ( $publisher <= ( $publisher_columns - $publisher_column_start ) ) {

        printXMLStartTag( "publisher", "2" );

        printXMLString( $publisher_organization[$publisher], "organizationName", "3" );

        if ( $publisher_address[$publisher] || $publisher_city[$publisher] || $publisher_state[$publisher] || $publisher_zipcode[$publisher] || $publisher_country[$publisher] ) {
            printXMLStartTag( "address", "3" );
            my $address = 0;
            my $publisher_address_lines;

            my @publisher_address_lines = split( /\|/, $publisher_address[$publisher] );
            while ( $address <= $#publisher_address_lines ) {

                printXMLString( $publisher_address_lines[$address], "deliveryPoint", "4" );
                $address = $address + 1;
            }
            printXMLString( $publisher_city[$publisher],    "city",               "4" );
            printXMLString( $publisher_state[$publisher],   "administrativeArea", "4" );
            printXMLString( $publisher_zipcode[$publisher], "postalCode",         "4" );
            printXMLString( $publisher_country[$publisher], "country",            "4" );
            printXMLEndTag( "address", "3" );
        }

        printXMLString( $publisher_phone[$publisher], "phone", "3", "voice", "phonetype" );
        printXMLString( $publisher_email[$publisher], "electronicMailAddress", "3" );
        printXMLString( $publisher_url[$publisher],   "onlineUrl",             "3" );

        printXMLEndTag( "publisher", "2" );

        $publisher = $publisher + 1;
    }

    #################################
    # Print DATASET METHODS SECTION #
    #################################

    if (   @dataset_methods_desc
        || @dataset_methods_citationID
        || @dataset_methods_protocolID
        || @dataset_quality_control_info
        || @dataset_methods_instrument
        || @dataset_sampling_desc
        || @dataset_studyext_desc
        || @sampling_sites_geographic_description )
    {

        printXMLStartTag( "methods", "2" );

        my $dataset_methods = 0;
        while ( $dataset_methods <= ( $dataset_methods_columns - $dataset_methods_column_start ) ) {

            printXMLStartTag( "methodStep", "3" );

            if (@dataset_methods_desc) {
                my @dataset_methods_desc_para = split( /\|/, $dataset_methods_desc[$dataset_methods] );
                my $methods_desc_para;
                printXMLStartTag( "description", "4" );
                foreach $methods_desc_para (@dataset_methods_desc_para) {

                    printXMLString( $methods_desc_para, "para", "5" );

                }
                printXMLEndTag( "description", "4" );
            }

            if (@dataset_methods_citationID) {
                my @dataset_methods_citations = split( /\|/, $dataset_methods_citationID[$dataset_methods] );
                my $methods_citation_row;
                foreach $methods_citation_row (@dataset_methods_citations) {
                    printXMLStartTag( "citation", "4" );

                    my $title = getStringValue( $WkS1->{Cells}[$methods_citation_row][1] );
                    printXMLString( $title, "title", "5" );

                    my $methods_citation_row1  = $methods_citation_row + 1;
                    my $methods_citation_row2  = $methods_citation_row + 2;
                    my $methods_citation_row3  = $methods_citation_row + 3;
                    my $methods_citation_row4  = $methods_citation_row + 4;
                    my $methods_citation_row5  = $methods_citation_row + 5;
                    my $methods_citation_row6  = $methods_citation_row + 6;
                    my $methods_citation_row7  = $methods_citation_row + 7;
                    my $methods_citation_row8  = $methods_citation_row + 8;
                    my $methods_citation_row9  = $methods_citation_row + 9;
                    my $methods_citation_row10 = $methods_citation_row + 10;
                    my $methods_citation_row11 = $methods_citation_row + 11;
                    my $methods_citation_row12 = $methods_citation_row + 12;
                    my $methods_citation_row13 = $methods_citation_row + 13;
                    my $methods_citation_row14 = $methods_citation_row + 14;
                    my $methods_citation_row15 = $methods_citation_row + 15;
                    my $methods_citation_row16 = $methods_citation_row + 16;
                    my $methods_citation_row17 = $methods_citation_row + 17;
                    my $methods_citation_row18 = $methods_citation_row + 18;
                    my $methods_citation_row19 = $methods_citation_row + 19;

                    my $author_rows_start   = $methods_citation_row1;
                    my $author_rows_end     = $methods_citation_row3;
                    my $author_column_start = 1;
                    my $author_columns      = getNumGroupColumns( $author_rows_start, $author_rows_end, 1, $WkS1 );
                    my @author_lastname     = getGroupedColumns( $author_columns, $methods_citation_row1, $WkS1, 1 );
                    my @author_firstname    = getGroupedColumns( $author_columns, $methods_citation_row2, $WkS1, 1 );
                    my @author_middlename   = getGroupedColumns( $author_columns, $methods_citation_row3, $WkS1, 1 );

                    my $author_creator = 0;

                    while ( $author_creator <= ( $author_columns - $author_column_start ) ) {

                        printXMLStartTag( "creator", "5" );

                        if ( $author_firstname[$author_creator] || $author_middlename[$author_creator] || $author_lastname[$author_creator] ) {
                            printXMLStartTag( "individualName", "6" );
                            printXMLString( $author_firstname[$author_creator],  "givenName", "7" );
                            printXMLString( $author_middlename[$author_creator], "givenName", "7" );
                            printXMLString( $author_lastname[$author_creator],   "surName",   "7" );
                            printXMLEndTag( "individualName", "6" );
                        }

                        printXMLEndTag( "creator", "5" );

                        $author_creator = $author_creator + 1;
                    }

                    my $publication_date = getStringValue( $WkS1->{Cells}[$methods_citation_row4][1] );
                    printXMLString( $publication_date, "pubDate", "5" );
                    my $citation_type = getStringValue( $WkS1->{Cells}[$methods_citation_row5][1] );

                    if ( $citation_type eq 'Article' ) {

                        my $journal        = getStringValue( $WkS1->{Cells}[$methods_citation_row6][1] );
                        my $journal_volume = getStringValue( $WkS1->{Cells}[$methods_citation_row7][1] );
                        my $journal_issue  = getStringValue( $WkS1->{Cells}[$methods_citation_row8][1] );
                        my $journal_pages  = getStringValue( $WkS1->{Cells}[$methods_citation_row9][1] );

                        printXMLStartTag( "article", "5" );
                        printXMLString( $journal,        "journal",   "6" );
                        printXMLString( $journal_volume, "volume",    "6" );
                        printXMLString( $journal_issue,  "issue",     "6" );
                        printXMLString( $journal_pages,  "pageRange", "6" );
                        printXMLEndTag( "article", "5" );

                    }

                    elsif ( $citation_type eq 'Book chapter' ) {

                        my $publisher   = getStringValue( $WkS1->{Cells}[$methods_citation_row6][1] );
                        my $pubplace    = getStringValue( $WkS1->{Cells}[$methods_citation_row7][1] );
                        my $edition     = getStringValue( $WkS1->{Cells}[$methods_citation_row8][1] );
                        my $total_pages = getStringValue( $WkS1->{Cells}[$methods_citation_row9][1] );
                        my $book_title  = getStringValue( $WkS1->{Cells}[$methods_citation_row10][1] );
                        my $page_range  = getStringValue( $WkS1->{Cells}[$methods_citation_row14][1] );

                        printXMLStartTag( "chapter", "5" );

                        if ($publisher) {
                            printXMLStartTag( "publisher", "6" );
                            printXMLString( $publisher, "organizationName", "7" );
                            printXMLEndTag( "publisher", "6" );
                        }

                        printXMLString( $pubplace,    "publicationPlace", "6" );
                        printXMLString( $edition,     "edition",          "6" );
                        printXMLString( $total_pages, "totalPages",       "6" );

                        my $editor_rows_start   = $methods_citation_row11;
                        my $editor_rows_end     = $methods_citation_row13;
                        my $editor_column_start = 1;
                        my $editor_columns      = getNumGroupColumns( $editor_rows_start, $editor_rows_end, 1, $WkS1 );
                        my @editor_lastname     = getGroupedColumns( $editor_columns, $methods_citation_row11, $WkS1, 1 );
                        my @editor_firstname    = getGroupedColumns( $editor_columns, $methods_citation_row12, $WkS1, 1 );
                        my @editor_middlename   = getGroupedColumns( $editor_columns, $methods_citation_row13, $WkS1, 1 );

                        if (@editor_lastname) {

                            printXMLStartTag( "editor", "5" );
                            my $editor = 0;

                            while ( $editor <= ( $editor_columns - $editor_column_start ) ) {

                                if ( $editor_firstname[$editor] || $editor_middlename[$editor] || $editor_lastname[$editor] ) {
                                    printXMLStartTag( "individualName", "6" );
                                    printXMLString( $editor_firstname[$editor],  "givenName", "7" );
                                    printXMLString( $editor_middlename[$editor], "givenName", "7" );
                                    printXMLString( $editor_lastname[$editor],   "surName",   "7" );
                                    printXMLEndTag( "individualName", "6" );
                                }

                                $editor = $editor + 1;
                            }

                            printXMLEndTag( "editor", "5" );
                        }

                        printXMLString( $book_title, "bookTitle", "5" );
                        printXMLString( $page_range, "pageRange", "5" );
                        printXMLEndTag( "chapter", "5" );

                    }

                    elsif ( $citation_type eq 'Book' ) {

                        my $publisher   = getStringValue( $WkS1->{Cells}[$methods_citation_row6][1] );
                        my $pubplace    = getStringValue( $WkS1->{Cells}[$methods_citation_row7][1] );
                        my $edition     = getStringValue( $WkS1->{Cells}[$methods_citation_row8][1] );
                        my $total_pages = getStringValue( $WkS1->{Cells}[$methods_citation_row9][1] );

                        printXMLStartTag( "book", "5" );

                        if ($publisher) {
                            printXMLStartTag( "publisher", "6" );
                            printXMLString( $publisher, "organizationName", "7" );
                            printXMLEndTag( "publisher", "6" );
                        }

                        printXMLString( $pubplace,    "publicationPlace", "6" );
                        printXMLString( $edition,     "edition",          "6" );
                        printXMLString( $total_pages, "totalPages",       "6" );

                        printXMLEndTag( "book", "5" );

                    }

                    elsif ( $citation_type eq 'Manuscript' ) {

                        my $institution = getStringValue( $WkS1->{Cells}[$methods_citation_row6][1] );
                        my $total_pages = getStringValue( $WkS1->{Cells}[$methods_citation_row7][1] );

                        printXMLStartTag( "manuscript", "5" );

                        if ($institution) {
                            printXMLStartTag( "institution", "6" );
                            printXMLString( $institution, "organizationName", "7" );
                            printXMLEndTag( "institution", "6" );
                        }

                        printXMLString( $total_pages, "totalPages", "6" );

                        printXMLEndTag( "manuscript", "5" );

                    }

                    elsif ( $citation_type eq 'Report' ) {

                        my $report_number = getStringValue( $WkS1->{Cells}[$methods_citation_row6][1] );
                        my $publisher     = getStringValue( $WkS1->{Cells}[$methods_citation_row7][1] );
                        my $pubplace      = getStringValue( $WkS1->{Cells}[$methods_citation_row8][1] );
                        my $total_pages   = getStringValue( $WkS1->{Cells}[$methods_citation_row9][1] );

                        printXMLStartTag( "report", "5" );
                        printXMLString( $report_number, "reportNumber", "6" );

                        if ($publisher) {
                            printXMLStartTag( "publisher", "6" );
                            printXMLString( $publisher, "organizationName", "7" );
                            printXMLEndTag( "publisher", "6" );
                        }

                        printXMLString( $pubplace,    "publicationPlace", "6" );
                        printXMLString( $total_pages, "totalPages",       "6" );

                        printXMLEndTag( "report", "5" );

                    }

                    elsif ( $citation_type eq 'Thesis' ) {

                        my $degree      = getStringValue( $WkS1->{Cells}[$methods_citation_row6][1] );
                        my $institution = getStringValue( $WkS1->{Cells}[$methods_citation_row7][1] );
                        my $total_pages = getStringValue( $WkS1->{Cells}[$methods_citation_row8][1] );

                        printXMLStartTag( "thesis", "5" );

                        printXMLString( $degree, "degree", "6" );

                        if ($institution) {
                            printXMLStartTag( "institution", "6" );
                            printXMLString( $institution, "organizationName", "7" );
                            printXMLEndTag( "institution", "6" );
                        }

                        printXMLString( $total_pages, "totalPages", "6" );

                        printXMLEndTag( "thesis", "5" );

                    }

                    elsif ( $citation_type eq 'Conference proceedings' ) {

                        my $publisher    = getStringValue( $WkS1->{Cells}[$methods_citation_row6][1] );
                        my $pubplace     = getStringValue( $WkS1->{Cells}[$methods_citation_row7][1] );
                        my $edition      = getStringValue( $WkS1->{Cells}[$methods_citation_row8][1] );
                        my $total_pages  = getStringValue( $WkS1->{Cells}[$methods_citation_row9][1] );
                        my $book_title   = getStringValue( $WkS1->{Cells}[$methods_citation_row10][1] );
                        my $page_range   = getStringValue( $WkS1->{Cells}[$methods_citation_row14][1] );
                        my $conf_name    = getStringValue( $WkS1->{Cells}[$methods_citation_row15][1] );
                        my $conf_date    = getStringValue( $WkS1->{Cells}[$methods_citation_row16][1] );
                        my $conf_city    = getStringValue( $WkS1->{Cells}[$methods_citation_row17][1] );
                        my $conf_state   = getStringValue( $WkS1->{Cells}[$methods_citation_row17][1] );
                        my $conf_country = getStringValue( $WkS1->{Cells}[$methods_citation_row17][1] );

                        printXMLStartTag( "conferenceProceedings", "5" );

                        if ($publisher) {
                            printXMLStartTag( "publisher", "6" );
                            printXMLString( $publisher, "organizationName", "7" );
                            printXMLEndTag( "publisher", "6" );
                        }

                        printXMLString( $pubplace,    "publicationPlace", "6" );
                        printXMLString( $edition,     "edition",          "6" );
                        printXMLString( $total_pages, "totalPages",       "6" );

                        my $editor_rows_start   = $methods_citation_row11;
                        my $editor_rows_end     = $methods_citation_row13;
                        my $editor_column_start = 1;
                        my $editor_columns      = getNumGroupColumns( $editor_rows_start, $editor_rows_end, 1, $WkS1 );
                        my @editor_lastname     = getGroupedColumns( $editor_columns, $methods_citation_row11, $WkS1, 1 );
                        my @editor_firstname    = getGroupedColumns( $editor_columns, $methods_citation_row12, $WkS1, 1 );
                        my @editor_middlename   = getGroupedColumns( $editor_columns, $methods_citation_row13, $WkS1, 1 );

                        if (@editor_lastname) {

                            printXMLStartTag( "editor", "5" );
                            my $editor = 0;

                            while ( $editor <= ( $editor_columns - $editor_column_start ) ) {

                                if ( $editor_firstname[$editor] || $editor_middlename[$editor] || $editor_lastname[$editor] ) {
                                    printXMLStartTag( "individualName", "6" );
                                    printXMLString( $editor_firstname[$editor],  "givenName", "7" );
                                    printXMLString( $editor_middlename[$editor], "givenName", "7" );
                                    printXMLString( $editor_lastname[$editor],   "surName",   "7" );
                                    printXMLEndTag( "individualName", "6" );
                                }

                                $editor = $editor + 1;
                            }

                            printXMLEndTag( "editor", "5" );
                        }

                        printXMLString( $book_title, "bookTitle",      "6" );
                        printXMLString( $page_range, "pageRange",      "6" );
                        printXMLString( $conf_name,  "conferenceName", "6" );
                        printXMLString( $conf_date,  "conferenceDate", "6" );
                        printXMLStartTag( "conferenceLocation", "6" );
                        printXMLString( $conf_city,    "city",               "7" );
                        printXMLString( $conf_state,   "administrativeArea", "7" );
                        printXMLString( $conf_country, "country",            "7" );
                        printXMLEndTag( "conferenceLocation",    "6" );
                        printXMLEndTag( "conferenceProceedings", "5" );

                    }

                    printXMLEndTag( "citation", "4" );
                }
            }

            if (@dataset_methods_protocolID) {
                my @dataset_methods_protocol = split( /\|/, $dataset_methods_protocolID[$dataset_methods] );
                my $methods_protocol_row;
                foreach $methods_protocol_row (@dataset_methods_protocol) {
                    printXMLStartTag( "protocol", "4" );

                    my $title = getStringValue( $WkS2->{Cells}[$methods_protocol_row][2] );
                    printXMLString( $title, "title", "5" );

                    my $methods_protocol_row1  = $methods_protocol_row + 1;
                    my $methods_protocol_row2  = $methods_protocol_row + 2;
                    my $methods_protocol_row3  = $methods_protocol_row + 3;
                    my $methods_protocol_row4  = $methods_protocol_row + 4;
                    my $methods_protocol_row5  = $methods_protocol_row + 5;
                    my $methods_protocol_row6  = $methods_protocol_row + 6;
                    my $methods_protocol_row7  = $methods_protocol_row + 7;
                    my $methods_protocol_row8  = $methods_protocol_row + 8;
                    my $methods_protocol_row9  = $methods_protocol_row + 9;
                    my $methods_protocol_row10 = $methods_protocol_row + 10;
                    my $methods_protocol_row11 = $methods_protocol_row + 11;
                    my $methods_protocol_row12 = $methods_protocol_row + 12;
                    my $methods_protocol_row13 = $methods_protocol_row + 13;
                    my $methods_protocol_row14 = $methods_protocol_row + 14;
                    my $methods_protocol_row15 = $methods_protocol_row + 15;
                    my $methods_protocol_row16 = $methods_protocol_row + 16;
                    my $methods_protocol_row17 = $methods_protocol_row + 17;
                    my $methods_protocol_row18 = $methods_protocol_row + 18;
                    my $methods_protocol_row19 = $methods_protocol_row + 19;

                    my $protocol_rows_start   = $methods_protocol_row1;
                    my $protocol_rows_end     = $methods_protocol_row14;
                    my $protocol_column_start = 2;
                    my $protocol_columns      = getNumGroupColumns( $protocol_rows_start, $protocol_rows_end, 2, $WkS2 );
                    my @protocol_salutation   = getGroupedColumns( $protocol_columns, $methods_protocol_row1, $WkS2, 2 );
                    my @protocol_firstname    = getGroupedColumns( $protocol_columns, $methods_protocol_row2, $WkS2, 2 );
                    my @protocol_lastname     = getGroupedColumns( $protocol_columns, $methods_protocol_row3, $WkS2, 2 );
                    my @protocol_organization = getGroupedColumns( $protocol_columns, $methods_protocol_row4, $WkS2, 2 );
                    my @protocol_position     = getGroupedColumns( $protocol_columns, $methods_protocol_row5, $WkS2, 2 );
                    my @protocol_address      = getGroupedColumns( $protocol_columns, $methods_protocol_row6, $WkS2, 2 );
                    my @protocol_city         = getGroupedColumns( $protocol_columns, $methods_protocol_row7, $WkS2, 2 );
                    my @protocol_state        = getGroupedColumns( $protocol_columns, $methods_protocol_row8, $WkS2, 2 );
                    my @protocol_zipcode      = getGroupedColumns( $protocol_columns, $methods_protocol_row9, $WkS2, 2 );
                    my @protocol_country      = getGroupedColumns( $protocol_columns, $methods_protocol_row10, $WkS2, 2 );
                    my @protocol_phone        = getGroupedColumns( $protocol_columns, $methods_protocol_row11, $WkS2, 2 );
                    my @protocol_fax          = getGroupedColumns( $protocol_columns, $methods_protocol_row12, $WkS2, 2 );
                    my @protocol_email        = getGroupedColumns( $protocol_columns, $methods_protocol_row13, $WkS2, 2 );
                    my @protocol_url          = getGroupedColumns( $protocol_columns, $methods_protocol_row14, $WkS2, 2 );

                    my $protocol = 0;

                    while ( $protocol <= ( $protocol_columns - $protocol_column_start ) ) {

                        printXMLStartTag( "creator", "5" );

                        if ( $protocol_salutation[$protocol] || $protocol_firstname[$protocol] || $protocol_lastname[$protocol] ) {
                            printXMLStartTag( "individualName", "6" );
                            printXMLString( $protocol_salutation[$protocol], "salutation", "7" );
                            printXMLString( $protocol_firstname[$protocol],  "givenName",  "7" );
                            printXMLString( $protocol_lastname[$protocol],   "surName",    "7" );
                            printXMLEndTag( "individualName", "6" );
                        }

                        if ( $protocol_organization[$protocol] ) {
                            printXMLString( $protocol_organization[$protocol], "organizationName", "6" );
                        }

                        if ( $protocol_position[$protocol] ) {
                            printXMLString( $protocol_position[$protocol], "positionName", "6" );
                        }

                        if (   $protocol_address[$protocol]
                            || $protocol_city[$protocol]
                            || $protocol_state[$protocol]
                            || $protocol_zipcode[$protocol]
                            || $protocol_country[$protocol] )
                        {
                            printXMLStartTag( "address", "6" );
                            my $address = 0;
                            my $protocol_address_lines;

                            my @protocol_address_lines = split( /\|/, $protocol_address[$protocol] );
                            while ( $address <= $#protocol_address_lines ) {

                                printXMLString( $protocol_address_lines[$address], "deliveryPoint", "7" );
                                $address = $address + 1;
                            }
                            printXMLString( $protocol_city[$protocol],    "city",               "7" );
                            printXMLString( $protocol_state[$protocol],   "administrativeArea", "7" );
                            printXMLString( $protocol_zipcode[$protocol], "postalCode",         "7" );
                            printXMLString( $protocol_country[$protocol], "country",            "7" );
                            printXMLEndTag( "address", "6" );
                        }

                        printXMLString( $protocol_phone[$protocol], "phone", "6", "voice", "phonetype" );
                        printXMLString( $protocol_fax[$protocol],   "phone", "6", "fax",   "phonetype" );
                        printXMLString( $protocol_email[$protocol], "electronicMailAddress", "6" );
                        printXMLString( $protocol_url[$protocol],   "onlineUrl",             "6" );

                        printXMLEndTag( "creator", "5" );

                        $protocol = $protocol + 1;
                    }
                    my $protocol_pubdate          = getStringValue( $WkS2->{Cells}[$methods_protocol_row15][2] );
                    my @protocol_abstract         = getArrayValueColumns( $methods_protocol_row16, 2, $WkS2 );
                    my @protocol_keywords         = getArrayValueColumns( $methods_protocol_row17, 2, $WkS2 );
                    my $protocol_url              = getStringValue( $WkS2->{Cells}[$methods_protocol_row18][2] );
                    my @protocol_procedural_steps = getArrayValueColumns( $methods_protocol_row19, 2, $WkS2 );

                    printXMLString( $protocol_pubdate, "pubDate", "5" );

                    if (@protocol_abstract) {
                        printXMLStartTag( "abstract", "5" );
                        my $protocol_abstract_para;
                        foreach $protocol_abstract_para (@protocol_abstract) {

                            printXMLString( $protocol_abstract_para, "para", "6" );

                        }
                        printXMLEndTag( "abstract", "5" );
                    }

                    if (@protocol_keywords) {
                        printXMLStartTag( "keywordSet", "5" );
                        my $protocol_keyword;
                        foreach $protocol_keyword (@protocol_keywords) {

                            printXMLString( $protocol_keyword, "keyword", "6" );

                        }
                        printXMLEndTag( "keywordSet", "5" );
                    }

                    if ($protocol_url) {
                        printXMLStartTag( "distribution", "5" );
                        printXMLStartTag( "online",       "6" );
                        printXMLString( $protocol_url, "url", "7" );
                        printXMLEndTag( "online",       "6" );
                        printXMLEndTag( "distribution", "5" );

                    }

                    if (@protocol_procedural_steps) {
                        my $protocol_procedural_step;
                        foreach $protocol_procedural_step (@protocol_procedural_steps) {
                            printXMLStartTag( "proceduralStep", "5" );
                            printXMLStartTag( "description",    "6" );
                            printXMLString( $protocol_procedural_step, "para", "7" );
                            printXMLEndTag( "description",    "6" );
                            printXMLEndTag( "proceduralStep", "5" );

                        }
                    }

                    printXMLEndTag( "protocol", "4" );
                }
            }

            if (@dataset_methods_instrument) {
                my @dataset_methods_instrument_para = split( /\|/, $dataset_methods_instrument[$dataset_methods] );
                my $methods_instrument_para;
                foreach $methods_instrument_para (@dataset_methods_instrument_para) {

                    printXMLString( $methods_instrument_para, "instrumentation", "4" );

                }
            }

            printXMLEndTag( "methodStep", "3" );

            $dataset_methods = $dataset_methods + 1;
        }

        if ( @dataset_studyext_desc || @dataset_sampling_desc || @sampling_sites_geographic_description ) {

            printXMLStartTag( "sampling", "3" );
            if (@dataset_studyext_desc) {
                printXMLStartTag( "studyExtent", "4" );

                my $studyext_desc = 0;

                while ( $studyext_desc <= $#dataset_studyext_desc ) {
                    printXMLStartTag( "description", "5" );
                    printXMLString( $dataset_studyext_desc[$studyext_desc], "para", "6" );
                    printXMLEndTag( "description", "5" );
                    $studyext_desc = $studyext_desc + 1;
                }

                printXMLEndTag( "studyExtent", "4" );
            }

            if (@dataset_sampling_desc) {
                printXMLStartTag( "samplingDescription", "4" );

                my $sampling_desc = 0;

                while ( $sampling_desc <= $#dataset_sampling_desc ) {
                    printXMLString( $dataset_sampling_desc[$sampling_desc], "para", "5" );
                    $sampling_desc = $sampling_desc + 1;
                }

                printXMLEndTag( "samplingDescription", "4" );
            }

            if (   @sampling_sites_geographic_description
                || @sampling_sites_latitude
                || @sampling_sites_longitude )
            {
                printXMLStartTag( "spatialSamplingUnits", "4" );

                my $sampling_sites_geo_coverage = 0;

                if (@sampling_sites_geographic_description) {
                    while ( $sampling_sites_geo_coverage <= ( $sampling_sites_geodesc_columns - $sampling_sites_geodesc_column_start ) ) {

                        printXMLStartTag( "coverage", "5" );

                        printXMLString( $sampling_sites_geographic_description[$sampling_sites_geo_coverage], "geographicDescription", "6" );

                        if (   $sampling_sites_latitude[$sampling_sites_geo_coverage]
                            || $sampling_sites_longitude[$sampling_sites_geo_coverage] )
                        {
                            printXMLStartTag( "boundingCoordinates", "6" );
                            printXMLString( $sampling_sites_longitude[$sampling_sites_geo_coverage], "westBoundingCoordinate",  "7" );
                            printXMLString( $sampling_sites_longitude[$sampling_sites_geo_coverage], "eastBoundingCoordinate",  "7" );
                            printXMLString( $sampling_sites_latitude[$sampling_sites_geo_coverage],  "northBoundingCoordinate", "7" );
                            printXMLString( $sampling_sites_latitude[$sampling_sites_geo_coverage],  "southBoundingCoordinate", "7" );
                            printXMLEndTag( "boundingCoordinates", "6" );

                        }

                        elsif ($sampling_sites_west_bounding_coord[$sampling_sites_geo_coverage]
                            || $sampling_sites_east_bounding_coord[$sampling_sites_geo_coverage]
                            || $sampling_sites_north_bounding_coord[$sampling_sites_geo_coverage]
                            || $sampling_sites_south_bounding_coord[$sampling_sites_geo_coverage] )
                        {
                            printXMLStartTag( "boundingCoordinates", "6" );
                            printXMLString( $sampling_sites_west_bounding_coord[$sampling_sites_geo_coverage],  "westBoundingCoordinate",  "7" );
                            printXMLString( $sampling_sites_east_bounding_coord[$sampling_sites_geo_coverage],  "eastBoundingCoordinate",  "7" );
                            printXMLString( $sampling_sites_north_bounding_coord[$sampling_sites_geo_coverage], "northBoundingCoordinate", "7" );
                            printXMLString( $sampling_sites_south_bounding_coord[$sampling_sites_geo_coverage], "southBoundingCoordinate", "7" );
                            printXMLEndTag( "boundingCoordinates", "6" );
                        }

                        printXMLEndTag( "coverage", "5" );

                        $sampling_sites_geo_coverage = $sampling_sites_geo_coverage + 1;
                    }
                }

                printXMLEndTag( "spatialSamplingUnits", "4" );
            }

            printXMLEndTag( "sampling", "3" );
        }

        if (@dataset_quality_control_info) {
            printXMLStartTag( "qualityControl", "3" );

            my $qualcontrol = 0;

            while ( $qualcontrol <= $#dataset_quality_control_info ) {
                printXMLStartTag( "description", "4" );
                printXMLString( $dataset_quality_control_info[$qualcontrol], "para", "5" );
                printXMLEndTag( "description", "4" );
                $qualcontrol = $qualcontrol + 1;
            }

            printXMLEndTag( "qualityControl", "3" );
        }

        printXMLEndTag( "methods", "2" );
    }

    ############
    percentDone;
    ############

    #########################
    # Print PROJECT section #
    #########################
    if (@research_project_number) {

        my $project_count = 0;
        my $print_relproj_start_tag;
        while ( $project_count <= $#research_project_number ) {
            my $project_start_row   = $research_project_number[$project_count];
            my $project_start_row1  = $project_start_row + 1;
            my $project_start_row2  = $project_start_row + 2;
            my $project_start_row3  = $project_start_row + 3;
            my $project_start_row4  = $project_start_row + 4;
            my $project_start_row5  = $project_start_row + 5;
            my $project_start_row6  = $project_start_row + 6;
            my $project_start_row7  = $project_start_row + 7;
            my $project_start_row8  = $project_start_row + 8;
            my $project_start_row9  = $project_start_row + 9;
            my $project_start_row10 = $project_start_row + 10;
            my $project_start_row11 = $project_start_row + 11;
            my $project_start_row12 = $project_start_row + 12;
            my $project_start_row13 = $project_start_row + 13;
            my $project_start_row14 = $project_start_row + 14;
            my $project_start_row15 = $project_start_row + 15;
            my $project_start_row16 = $project_start_row + 16;
            my $project_start_row17 = $project_start_row + 17;
            my $project_start_row18 = $project_start_row + 18;
            my $project_start_row19 = $project_start_row + 19;
            my $project_start_row20 = $project_start_row + 20;
            my $project_start_row21 = $project_start_row + 21;
            my $project_start_row22 = $project_start_row + 22;
            my $project_start_row23 = $project_start_row + 23;

            my $project_indent1 = $project_count + 1;
            my $project_indent2 = $project_count + 2;
            my $project_indent3 = $project_count + 3;
            my $project_indent4 = $project_count + 4;
            my $project_indent5 = $project_count + 5;
            my $project_indent6 = $project_count + 6;
            my $project_indent7 = $project_count + 7;
            my $project_indent8 = $project_count + 8;
            my $project_indent9 = $project_count + 9;

            my $research_project_ID    = getStringValue( $WkS3->{Cells}[$project_start_row][2] );
            my $research_project_title = getStringValue( $WkS3->{Cells}[$project_start_row1][2] );

            my $projpers_rows_start                   = $project_start_row2;
            my $projpers_rows_end                     = $project_start_row16;
            my $projpers_column_start                 = 2;
            my $projpers_columns                      = getNumGroupColumns( $projpers_rows_start, $projpers_rows_end, 2, $WkS3 );
            my @research_project_firstname            = getGroupedColumns( $projpers_columns, $project_start_row2, $WkS3, 2 );
            my @research_project_lastname             = getGroupedColumns( $projpers_columns, $project_start_row3, $WkS3, 2 );
            my @research_project_role                 = getGroupedColumns( $projpers_columns, $project_start_row4, $WkS3, 2 );
            my @research_project_organization         = getGroupedColumns( $projpers_columns, $project_start_row5, $WkS3, 2 );
            my @research_project_position             = getGroupedColumns( $projpers_columns, $project_start_row6, $WkS3, 2 );
            my @research_project_address              = getGroupedColumns( $projpers_columns, $project_start_row7, $WkS3, 2 );
            my @research_project_city                 = getGroupedColumns( $projpers_columns, $project_start_row8, $WkS3, 2 );
            my @research_project_state                = getGroupedColumns( $projpers_columns, $project_start_row9, $WkS3, 2 );
            my @research_project_zipcode              = getGroupedColumns( $projpers_columns, $project_start_row10, $WkS3, 2 );
            my @research_project_country              = getGroupedColumns( $projpers_columns, $project_start_row11, $WkS3, 2 );
            my @research_project_phone                = getGroupedColumns( $projpers_columns, $project_start_row12, $WkS3, 2 );
            my @research_project_fax                  = getGroupedColumns( $projpers_columns, $project_start_row13, $WkS3, 2 );
            my @research_project_email                = getGroupedColumns( $projpers_columns, $project_start_row14, $WkS3, 2 );
            my $research_project_url                  = getStringValue( $WkS3->{Cells}[$project_start_row15][2] );
            my $research_project_geographic_desc      = getStringValue( $WkS3->{Cells}[$project_start_row16][2] );
            my $research_project_west_bounding_coord  = getStringValue( $WkS3->{Cells}[$project_start_row17][2] );
            my $research_project_east_bounding_coord  = getStringValue( $WkS3->{Cells}[$project_start_row18][2] );
            my $research_project_north_bounding_coord = getStringValue( $WkS3->{Cells}[$project_start_row19][2] );
            my $research_project_south_bounding_coord = getStringValue( $WkS3->{Cells}[$project_start_row20][2] );
            my @research_project_temporal_coverage    = getArrayValue( $WkS3->{Cells}[$project_start_row21][2] );
            my @research_project_abstract             = getArrayValueColumns( $project_start_row22, 2, $WkS3 );
            my @research_project_funding              = getArrayValue( $WkS3->{Cells}[$project_start_row23][2] );

            if ( $project_count == 0 ) {

                if ($research_project_ID) {
                    printXMLStartTag( "project", "$project_indent2", $research_project_ID );
                }
                else {
                    printXMLStartTag( "project", "$project_indent2" );
                }
            }
            elsif ( $print_relproj_start_tag eq 'yes' ) {

                if ($research_project_ID) {
                    printXMLStartTag( "relatedProject", "$project_indent2", $research_project_ID );
                }
                else {
                    printXMLStartTag( "relatedProject", "$project_indent2" );
                }
            }

            printXMLString( $research_project_title, "title", "$project_indent3" );

            my $research_project_pers = 0;

            while ( $research_project_pers <= ( $projpers_columns - $projpers_column_start ) ) {
                printXMLStartTag( "personnel", "$project_indent4" );
                if ( $research_project_firstname[$research_project_pers] || $research_project_lastname[$research_project_pers] ) {
                    printXMLStartTag( "individualName", "$project_indent5" );
                    printXMLString( $research_project_firstname[$research_project_pers], "givenName", "$project_indent6" );
                    printXMLString( $research_project_lastname[$research_project_pers],  "surName",   "$project_indent6" );
                    printXMLEndTag( "individualName", "$project_indent5" );
                }

                if ( $research_project_organization[$research_project_pers] ) {
                    printXMLString( $research_project_organization[$research_project_pers], "organizationName", "$project_indent5" );
                }

                if ( $research_project_position[$research_project_pers] ) {
                    printXMLString( $research_project_position[$research_project_pers], "positionName", "$project_indent5" );
                }

                if (   $research_project_address[$research_project_pers]
                    || $research_project_city[$research_project_pers]
                    || $research_project_state[$research_project_pers]
                    || $research_project_zipcode[$research_project_pers]
                    || $research_project_country[$research_project_pers] )
                {
                    printXMLStartTag( "address", "$project_indent5" );

                    my $address = 0;
                    my $research_project_address_lines;

                    my @research_project_address_lines =
                      split( /\|/, $research_project_address[$research_project_pers] );
                    while ( $address <= $#research_project_address_lines ) {

                        printXMLString( $research_project_address_lines[$address], "deliveryPoint", "$project_indent6" );
                        $address = $address + 1;
                    }

                    printXMLString( $research_project_city[$research_project_pers],    "city",               "$project_indent6" );
                    printXMLString( $research_project_state[$research_project_pers],   "administrativeArea", "$project_indent6" );
                    printXMLString( $research_project_zipcode[$research_project_pers], "postalCode",         "$project_indent6" );
                    printXMLString( $research_project_country[$research_project_pers], "country",            "$project_indent6" );
                    printXMLEndTag( "address", "$project_indent5" );

                }

                printXMLString( $research_project_phone[$research_project_pers], "phone", "$project_indent5", "voice", "phonetype" );
                printXMLString( $research_project_fax[$research_project_pers],   "phone", "$project_indent5", "fax",   "phonetype" );
                printXMLString( $research_project_email[$research_project_pers], "electronicMailAddress", "$project_indent5" );

                printXMLString( $research_project_role[$research_project_pers], "role", "$project_indent5" );

                printXMLEndTag( "personnel", "$project_indent4" );

                $research_project_pers = $research_project_pers + 1;
            }

            if (@research_project_abstract) {
                printXMLStartTag( "abstract", "$project_indent4" );
                my $abstract = 0;
                while ( $abstract <= $#research_project_abstract ) {
                    printXMLString( $research_project_abstract[$abstract], "para", "$project_indent5" );
                    $abstract = $abstract + 1;
                }
                printXMLEndTag( "abstract", "$project_indent4" );
            }

            if (@research_project_funding) {
                printXMLStartTag( "funding", "$project_indent4" );
                my $funding = 0;
                while ( $funding <= $#research_project_funding ) {
                    printXMLString( $research_project_funding[$funding], "para", "$project_indent5" );
                    $funding = $funding + 1;
                }
                printXMLEndTag( "funding", "$project_indent4" );
            }

            if (   $research_project_geographic_desc
                || $research_project_west_bounding_coord
                || $research_project_east_bounding_coord
                || $research_project_north_bounding_coord
                || $research_project_south_bounding_coord
                || @research_project_temporal_coverage )
            {

                printXMLStartTag( "studyAreaDescription", "$project_indent4" );
                printXMLStartTag( "coverage",             "$project_indent4" );

                if (   $research_project_geographic_desc
                    || $research_project_west_bounding_coord
                    || $research_project_east_bounding_coord
                    || $research_project_north_bounding_coord
                    || $research_project_south_bounding_coord )
                {

                    printXMLStartTag( "geographicCoverage", "$project_indent5" );

                    if ($research_project_geographic_desc) {
                        printXMLString( $research_project_geographic_desc, "geographicDescription", "$project_indent6" );
                    }

                    if (   $research_project_west_bounding_coord
                        || $research_project_east_bounding_coord
                        || $research_project_north_bounding_coord
                        || $research_project_south_bounding_coord )
                    {
                        printXMLStartTag( "boundingCoordinates", "$project_indent6" );
                        printXMLString( $research_project_west_bounding_coord,  "westBoundingCoordinate",  "$project_indent7" );
                        printXMLString( $research_project_east_bounding_coord,  "eastBoundingCoordinate",  "$project_indent7" );
                        printXMLString( $research_project_north_bounding_coord, "northBoundingCoordinate", "$project_indent7" );
                        printXMLString( $research_project_south_bounding_coord, "southBoundingCoordinate", "$project_indent7" );
                        printXMLEndTag( "boundingCoordinates", "$project_indent6" );
                    }

                    printXMLEndTag( "geographicCoverage", "$project_indent5" );

                }

                if (@research_project_temporal_coverage) {
                    printXMLStartTag( "temporalCoverage", "$project_indent5" );
                    printXMLStartTag( "rangeOfDates",     "$project_indent6" );

                    if ( $research_project_temporal_coverage[0] ) {
                        printXMLStartTag( "beginDate", "$project_indent7" );
                        printXMLString( $research_project_temporal_coverage[0], "calendarDate", "$project_indent8" );
                        printXMLEndTag( "beginDate", "$project_indent7" );
                    }

                    if ( $research_project_temporal_coverage[1] ) {
                        printXMLStartTag( "endDate", "$project_indent7" );
                        printXMLString( $research_project_temporal_coverage[1], "calendarDate", "$project_indent8" );
                        printXMLEndTag( "endDate", "$project_indent7" );
                    }

                    printXMLEndTag( "rangeOfDates",     "$project_indent6" );
                    printXMLEndTag( "temporalCoverage", "$project_indent5" );
                }

                printXMLEndTag( "coverage",             "$project_indent4" );
                printXMLEndTag( "studyAreaDescription", "$project_indent4" );
            }
            if ( $project_count < $#research_project_number && $project_count == 0 ) {
                $print_relproj_start_tag = "yes";
            }
            elsif ( $project_count < $#research_project_number && $project_count > 0 ) {
                printXMLEndTag( "relatedProject", "$project_indent2" );
                $print_relproj_start_tag = "yes";
            }
            elsif ( $project_count == $#research_project_number && $project_count > 0 ) {
                printXMLEndTag( "relatedProject", "$project_indent2" );
                printXMLEndTag( "project",        "2" );
            }
            else {
                printXMLEndTag( "project", "2" );

            }
            $project_count = $project_count + 1;

        }

    }

    ########################
    # Print ACCESS SECTION #
    ########################

    print XML "$indent$indent" . "<access " . "$dataset_access_authentication_info" . ">\n";

    my $access = 0;

    while ( $access <= $#dataset_principal_access_info ) {

        if ( @dataset_principal_permission_info && @dataset_principal_access_info ) {

            printXMLStartTag( "allow", "3" );

            printXMLString( $dataset_principal_access_info[$access],     "principal",  "4" );
            printXMLString( $dataset_principal_permission_info[$access], "permission", "4" );

            printXMLEndTag( "allow", "3" );

        }
        $access = $access + 1;
    }
    printXMLEndTag( "access", "2" );

    ############################
    # Print DATATABLE section  #
    ############################

    printXMLStartTag( "dataTable", "2" );

    if (@data_entity_name) {
        my $entity_name = 0;

        while ( $entity_name <= $#data_entity_name ) {

            printXMLString( $data_entity_name[$entity_name], "entityName",        "3" );
            printXMLString( $data_entity_desc[$entity_name], "entityDescription", "3" );
            $entity_name = $entity_name + 1;
        }

        printXMLStartTag( "physical", "3" );
        printXMLString( $data_object_name, "objectName", "4" );

        if (   $num_header_lines
            || $num_data_records
            || $data_attribute_orientation
            || $data_field_delimiter
            || $data_external_format )
        {

            printXMLStartTag( "dataFormat", "4" );
            if (   $num_header_lines
                || $data_attribute_orientation
                || $data_field_delimiter )
            {

                #printXMLString( $num_data_records,           "size",      "4", "rows", "units" );

                printXMLStartTag( "textFormat", "5" );
                printXMLString( $num_header_lines,           "numHeaderLines",       "6" );
                printXMLString( $data_attribute_orientation, "attributeOrientation", "6" );

                if ($data_field_delimiter) {
                    printXMLStartTag( "simpleDelimited", "6" );
                    printXMLString( $data_field_delimiter, "fieldDelimiter", "7" );
                    printXMLEndTag( "simpleDelimited", "6" );
                }

                printXMLEndTag( "textFormat", "5" );

            }

            elsif ($data_external_format) {
                printXMLStartTag( "externallyDefinedFormat", "5" );
                printXMLString( $data_external_format, "formatName", "6" );
                printXMLEndTag( "externallyDefinedFormat", "5" );

            }

            printXMLEndTag( "dataFormat", "4" );
        }

        printXMLEndTag( "physical", "3" );

        ############
        percentDone;
        ############

        ############################
        # Print ATTRIBUTES section #
        ############################

        if (@attribute_name) {
            printXMLStartTag( "attributeList", "2" );

            my $attribute_num   = 1;
            my $attribute_count = 0;

            while ( $attribute_count <= $#attribute_name ) {

                printXMLStartTag( "attribute", "3", "att.$attribute_num" );
                printXMLString( $attribute_name[$attribute_count],       "attributeName",       "4" );
                printXMLString( $attribute_label[$attribute_count],      "attributeLabel",      "4" );
                printXMLString( $attribute_definition[$attribute_count], "attributeDefinition", "4" );
                printXMLString( $variable_type[$attribute_count],        "storageType",         "4" );

                if ( $measurement_scale[$attribute_count] eq 'nominal' ) {
                    printXMLStartTag( "measurementScale", "4" );
                    printXMLStartTag( "nominal",          "5" );
                    printXMLStartTag( "nonNumericDomain", "6" );

                    if ( $codeset_name[$attribute_count] ) {
                        printXMLStartTag( "enumeratedDomain", "7" );

                        my @codes = split( /\|/, $codeset_name[$attribute_count] );
                        my $pair;
                        foreach $pair (@codes) {
                            my @codeset = split( /\=/, $pair );
                            printXMLStartTag( "codeDefinition", "8" );
                            printXMLString( $codeset[0], "code",       "9" );
                            printXMLString( $codeset[1], "definition", "9" );
                            printXMLEndTag( "codeDefinition", "8" );
                        }

                        printXMLEndTag( "enumeratedDomain", "7" );
                    }
                    else {
                        printXMLStartTag( "textDomain", "7" );
                        printXMLString( $attribute_definition[$attribute_count], "definition", "9" );
                        printXMLEndTag( "textDomain", "7" );
                    }

                    printXMLEndTag( "nonNumericDomain", "6" );
                    printXMLEndTag( "nominal",          "5" );
                    printXMLEndTag( "measurementScale", "4" );

                }
                if ( $measurement_scale[$attribute_count] eq 'ordinal' ) {
                    printXMLStartTag( "measurementScale", "4" );
                    printXMLStartTag( "ordinal",          "5" );
                    printXMLStartTag( "nonNumericDomain", "6" );

                    if ( $codeset_name[$attribute_count] ) {
                        printXMLStartTag( "enumeratedDomain", "7" );

                        my @codes = split( /\|/, $codeset_name[$attribute_count] );
                        my $pair;
                        foreach $pair (@codes) {
                            my @codeset = split( /\=/, $pair );
                            printXMLStartTag( "codeDefinition", "8" );
                            printXMLString( $codeset[0], "code",       "9" );
                            printXMLString( $codeset[1], "definition", "9" );
                            printXMLEndTag( "codeDefinition", "8" );
                        }

                        printXMLEndTag( "enumeratedDomain", "7" );
                    }
                    else {
                        printXMLStartTag( "textDomain", "7" );
                        printXMLString( $attribute_definition[$attribute_count], "definition", "9" );
                        printXMLEndTag( "textDomain", "7" );
                    }
                    printXMLEndTag( "nonNumericDomain", "6" );
                    printXMLEndTag( "ordinal",          "5" );
                    printXMLEndTag( "measurementScale", "4" );
                }
                elsif ( $measurement_scale[$attribute_count] eq 'datetime' ) {
                    printXMLStartTag( "measurementScale", "4" );
                    printXMLStartTag( "datetime",         "5" );
                    printXMLString( $date_time_format[$attribute_count], "formatString",      "6" );
                    printXMLString( $precision[$attribute_count],        "dateTimePrecision", "6" );
                    printXMLStartTag( "dateTimeDomain", "6" );
                    printXMLStartTag( "bounds",         "7" );
                    if ( @date_time_min && @date_time_max ) {
                        printXMLString( $date_time_min[$attribute_count], "minimum", "8", "false", "exclusive" );
                        printXMLString( $date_time_max[$attribute_count], "maximum", "8", "false", "exclusive" );
                    }
                    printXMLEndTag( "bounds",           "7" );
                    printXMLEndTag( "dateTimeDomain",   "6" );
                    printXMLEndTag( "datetime",         "5" );
                    printXMLEndTag( "measurementScale", "4" );
                }
                elsif ( $measurement_scale[$attribute_count] eq 'interval' ) {
                    printXMLStartTag( "measurementScale", "4" );
                    printXMLStartTag( "interval",         "5" );
                    printXMLStartTag( "unit",             "6" );

                    if ( $custom_or_eml[$attribute_count] eq 'EML' ) {
                        printXMLString( $units[$attribute_count], "standardUnit", "7" );
                    }

                    elsif ( $custom_or_eml[$attribute_count] eq 'CUSTOM' ) {
                        my $unit;
                        my $repeat = 0;
                        if (@custom_unit_list) {
                            foreach $unit (@custom_unit_list) {
                                if ( $unit eq $units[$attribute_count] ) {
                                    $repeat = "yes";
                                }
                            }
                        }
                        if ( $repeat eq "yes" ) {
                        }
                        else {
                            my $custom_unit_stmml;
                            $custom_unit_stmml =
                                "<stmml:unit name=\""
                              . "$units[$attribute_count]"
                              . "\" unitType=\""
                              . "$custom_unitType[$attribute_count]"
                              . "\" id=\""
                              . "$custom_unitID[$attribute_count]"
                              . "\" parentSI=\""
                              . "$custom_unitParentSI[$attribute_count]"
                              . "\"  multiplierToSI=\""
                              . "$custom_unitMultiplierToSI[$attribute_count]" . "\">";
                            push( @custom_unit_stmml_tag, $custom_unit_stmml );
                            push( @custom_unit_list,      $units[$attribute_count] );
                        }
                        printXMLString( $units[$attribute_count], "customUnit", "7" );
                    }
                    else {
                    }

                    printXMLEndTag( "unit", "6" );
                    printXMLString( $precision[$attribute_count], "precision", "7" );
                    printXMLStartTag( "numericDomain", "7" );
                    printXMLString( $number_type[$attribute_count], "numberType", "8" );
                    printXMLEndTag( "numericDomain",    "7" );
                    printXMLEndTag( "interval",         "5" );
                    printXMLEndTag( "measurementScale", "4" );

                }
                elsif ( $measurement_scale[$attribute_count] eq 'ratio' ) {
                    printXMLStartTag( "measurementScale", "4" );
                    printXMLStartTag( "ratio",            "5" );
                    printXMLStartTag( "unit",             "6" );

                    if ( $custom_or_eml[$attribute_count] eq 'EML' ) {
                        printXMLString( $units[$attribute_count], "standardUnit", "7" );
                    }

                    elsif ( $custom_or_eml[$attribute_count] eq 'CUSTOM' ) {
                        my $unit;
                        my $repeat = 0;
                        if (@custom_unit_list) {
                            foreach $unit (@custom_unit_list) {
                                if ( $unit eq $units[$attribute_count] ) {
                                    $repeat = "yes";
                                }
                            }
                        }
                        if ( $repeat eq "yes" ) {
                        }
                        else {
                            my $custom_unit_stmml;
                            my $custom_unit_stmml_desc;

                            if (   !$custom_unitParentSI[$attribute_count]
                                && !$custom_unitMultiplierToSI[$attribute_count] )
                            {

                                $custom_unit_stmml =
                                    "<stmml:unit name=\""
                                  . "$units[$attribute_count]"
                                  . "\" unitType=\""
                                  . "$custom_unitType[$attribute_count]"
                                  . "\" id=\""
                                  . "$custom_unitID[$attribute_count]" . "\">";
                            }
                            elsif ( !$custom_unitParentSI[$attribute_count]
                                && $custom_unitMultiplierToSI[$attribute_count] )
                            {

                                $custom_unit_stmml =
                                    "<stmml:unit name=\""
                                  . "$units[$attribute_count]"
                                  . "\" unitType=\""
                                  . "$custom_unitType[$attribute_count]"
                                  . "\" id=\""
                                  . "$custom_unitID[$attribute_count]"
                                  . "\"  multiplierToSI=\""
                                  . "$custom_unitMultiplierToSI[$attribute_count]" . "\">";
                            }
                            else {
                                $custom_unit_stmml =
                                    "<stmml:unit name=\""
                                  . "$units[$attribute_count]"
                                  . "\" unitType=\""
                                  . "$custom_unitType[$attribute_count]"
                                  . "\" id=\""
                                  . "$custom_unitID[$attribute_count]"
                                  . "\" parentSI=\""
                                  . "$custom_unitParentSI[$attribute_count]"
                                  . "\"  multiplierToSI=\""
                                  . "$custom_unitMultiplierToSI[$attribute_count]" . "\">";
                            }
                            $custom_unit_stmml_desc = "<stmml:description>" . "$custom_unitDesc[$attribute_count]" . "</stmml:description>";

                            push( @custom_unit_stmml_tag,      $custom_unit_stmml );
                            push( @custom_unit_stmml_desc_tag, $custom_unit_stmml_desc );
                            push( @custom_unit_list,           $units[$attribute_count] );
                        }
                        printXMLString( $units[$attribute_count], "customUnit", "7" );
                    }
                    else {
                    }

                    printXMLEndTag( "unit", "6" );
                    printXMLString( $precision[$attribute_count], "precision", "7" );
                    printXMLStartTag( "numericDomain", "7" );
                    printXMLString( $number_type[$attribute_count], "numberType", "8" );
                    printXMLEndTag( "numericDomain",    "7" );
                    printXMLEndTag( "ratio",            "5" );
                    printXMLEndTag( "measurementScale", "4" );

                }

                else {
                }

                if ( $missing_value_code[$attribute_count] ) {
                    printXMLStartTag( "missingValueCode", "4" );
                    printXMLString( $missing_value_code[$attribute_count],        "code",            "5" );
                    printXMLString( $missing_value_explanation[$attribute_count], "codeExplanation", "5" );
                    printXMLEndTag( "missingValueCode", "4" );
                }

                if ( $calculations[$attribute_count] ) {
                    my $calculations_text = "Calculations: " . $calculations[$attribute_count];
                    printXMLStartTag( "method",      "4" );
                    printXMLStartTag( "methodStep",  "5" );
                    printXMLStartTag( "description", "6" );
                    printXMLString( $calculations_text, "para", "7" );
                    printXMLEndTag( "description", "6" );
                    printXMLEndTag( "methodStep",  "5" );
                    printXMLEndTag( "method",      "4" );
                }

                printXMLEndTag( "attribute", "3" );
                $attribute_num   = $attribute_num + 1;
                $attribute_count = $attribute_count + 1;

            }

            printXMLEndTag( "attributeList", "2" );
        }
    }
    printXMLString( $num_data_records, "numberOfRecords", "3" );

    printXMLEndTag( "dataTable", "2" );

    printXMLEndTag( "dataset", "1" );

    #####################################
    # Print ADDITIONAL METADATA section #
    #####################################

    if (@custom_unit_stmml_tag) {

        printXMLStartTag( "additionalMetadata", "0" );
        print XML "$indent" . "<stmml:unitList xmlns:stmml=\"http://www.xml-cml.org/schema/stmml\"\n";
        print XML "$indent" . "xsi:schemaLocation=\"http://www.xml-cml.org/schema/stmml $stmml\">\n";

        my $custom_stmml;
        my $custom_stmml_desc;
        my $custom_stmml_count = 0;

        foreach $custom_stmml (@custom_unit_stmml_tag) {

            print XML "$indent$indent" . "$custom_stmml\n";
            print XML "$indent$indent$indent" . "$custom_unit_stmml_desc_tag[$custom_stmml_count]\n";

            print XML "$indent$indent" . "</stmml:unit>\n";
            $custom_stmml_count = $custom_stmml_count + 1;
        }

        print XML "$indent" . "</stmml:unitList>\n";
        printXMLEndTag( "additionalMetadata", "0" );
    }

    my $addM_col      = 2;
    my $addM_row      = 149;
    my $addM_tag_col  = 0;
    my $addM_continue = "yes";
    my @previous_row_tags;
    my @previous_row;
    my @next_row;
    my $count = 0;
    my $grouped_rows;

    while ($WkS0->{Cells}[$addM_row][$addM_tag_col]
        && $addM_continue eq 'yes' )
    {

        my @row = getArrayValueColumns( $addM_row, $addM_col, $WkS0 );
        my @row_tags = getArrayValue( $WkS0->{Cells}[$addM_row][$addM_tag_col] );
        @previous_row_tags = getArrayValue( $WkS0->{Cells}[ $addM_row - 1 ][$addM_tag_col] );
        @previous_row      = getArrayValue( $WkS0->{Cells}[ $addM_row - 1 ][$addM_col] );
        @next_row          = getArrayValue( $WkS0->{Cells}[ $addM_row + 1 ][$addM_col] );
        my @next_row_tags = getArrayValue( $WkS0->{Cells}[ $addM_row + 1 ][$addM_tag_col] );

        if ( @row_tags && @row ) {

            if ( $count == 0 ) {
                printXMLStartTag( "additionalMetadata", "0" );
            }
            my $start_tag = 0;
            while ( ( $start_tag < $#row_tags ) && @row ) {
                if ( ( $row_tags[$start_tag] eq $previous_row_tags[$start_tag] ) && $count > 0 ) {
                    $start_tag = $start_tag + 1;
                }
                else {
                    printXMLStartTag( $row_tags[$start_tag], $start_tag );
                    $start_tag = $start_tag + 1;
                }
            }

            my $row_value;
            foreach $row_value (@row) {
                printXMLString( $row_value, $row_tags[$#row_tags], $#row_tags );

            }

            $count = $count + 1;

            my $end_tag = $#row_tags - 1;

            while ( $end_tag >= 0 ) {

                if ( $row_tags[$end_tag] eq $next_row_tags[$end_tag] ) {

                    my $group_row        = $addM_row + 1;
                    my @grouped_row_tags = getArrayValue( $WkS0->{Cells}[$group_row][$addM_tag_col] );
                    my @grouped_row;
                    my $grouped_rows;

                    while ( $row_tags[$end_tag] eq $grouped_row_tags[$end_tag] ) {
                        @grouped_row_tags = getArrayValue( $WkS0->{Cells}[$group_row][$addM_tag_col] );
                        @grouped_row      = getArrayValue( $WkS0->{Cells}[ $group_row - 1 ][$addM_col] );

                        if ( @grouped_row && ( $group_row > ( $addM_row + 1 ) ) ) {
                            $grouped_rows .= "more";
                        }
                        else {
                            $grouped_rows .= "end";
                        }
                        $group_row = $group_row + 1;
                    }

                    if ( $grouped_rows =~ /more/ ) {
                        $end_tag = $end_tag - 1;
                    }
                    else {
                        printXMLEndTag( $row_tags[$end_tag], $end_tag );
                        $end_tag = $end_tag - 1;
                    }

                }

                else {
                    printXMLEndTag( $row_tags[$end_tag], $end_tag );
                    $end_tag = $end_tag - 1;
                }
            }

            $addM_row = $addM_row + 1;
        }
        elsif (@next_row_tags) {
            $addM_continue = "yes";
            $addM_row      = $addM_row + 1;

        }
        else {
            $addM_continue = "no";
        }

    }

    if ( $count > 0 ) {
        printXMLEndTag( "additionalMetadata", "0" );
    }

    print XML "</eml:eml>\n";

    close(XML);

################################################################################################
# EML Schema Validation   (SAX validation using XML::Xerces)                                   #
#                                                                                              #
# Slightly modified portion of the validate.pl sample distributed with the XML::Xerces module  #
# Validates the newly created eml file against the schema specified in the EML file.           #
# Error messages are displayed in the application's log and recorded in an error.log file      #
# This is the only portion of the file that uses XML::Xerces, IO::Handle, and Getopt::Long;    #
################################################################################################

    if ( $validation_checkbox eq 'yes' ) {

        $lb_out->insert( "end", "   Done!  Validating against schema..." );

        package XMLErrorHandler;
        use strict;
        use vars qw(@ISA);
        @ISA = qw(XML::Xerces::PerlErrorHandler);
        use vars qw($cwd);
        use vars qw($error_type);

        sub cwd_directory_die {
            my $cwd = $_[0];
            $lb_out->insert( "end", "  " );
            $lb_out->insert( "end", ":-O  Can't write errors to error.log in the current working directory." );
            $lb_out->insert( "end", "       ($cwd)." );
            $lb_out->insert( "end", "       Please verify that this directory exists and that you can write to the directory." );
        }

        sub warning {
            my $time    = localtime;
            my $line    = $_[1]->getLineNumber;
            my $column  = $_[1]->getColumnNumber;
            my $message = $_[1]->getMessage;
            my $warning = "EML Validation Warning - " . "$main::FILE" . ", line=" . "$line" . ", " . " $message";
            $lb_out->insert( "end", "     " . "$warning" );
            open( ERROR, ">>error.log" ) or &cwd_directory_die($cwd);
            print ERROR "$time" . "  " . "$warning\n";
            close(ERROR);
            $error_type = "Warnings";
        }

        sub error {
            my $time    = localtime;
            my $line    = $_[1]->getLineNumber;
            my $column  = $_[1]->getColumnNumber;
            my $message = $_[1]->getMessage;
            my $error   = "EML Validation Error - " . "$main::FILE" . ", line=" . "$line" . ", " . " $message";
            $lb_out->insert( "end", "     " . "$error" );
            open( ERROR, ">>error.log" ) or &cwd_directory_die($cwd);
            print ERROR "$time" . "  " . "$error\n";
            close(ERROR);
            $error_type = "Errors";
        }

        sub fatal_error {
            my $time    = localtime;
            my $line    = $_[1]->getLineNumber;
            my $column  = $_[1]->getColumnNumber;
            my $message = $_[1]->getMessage;
            my $error   = "EML Validation Error - " . "$main::FILE" . ", line=" . "$line" . ", " . " $message";
            $lb_out->insert( "end", "     " . "$error" );
            open( ERROR, ">>error.log" ) or &cwd_directory_die($cwd);
            print ERROR "$time" . "  " . "$error\n";
            close(ERROR);
            $error_type = "Errors";
        }

        1;

        package main;

        my $error_type;
        my $cwd = getcwd;
        $main::FILE = $eml_file;

        my $parser = XML::Xerces::XMLReaderFactory::createXMLReader();
        $parser->setErrorHandler( XMLErrorHandler->new );

        my $contentHandler = new XML::Xerces::PerlContentHandler();
        $parser->setContentHandler($contentHandler);

        # handle the optional features
        eval { $parser->setFeature( "$XML::Xerces::XMLUni::fgXercesSchema", 1 ); };
        XML::Xerces::error($@) if $@;

        ############
        percentDone;
        ############

        # and the required features
        eval {
            $parser->setFeature( "$XML::Xerces::XMLUni::fgXercesContinueAfterFatalError", 1 );
            $parser->setFeature( "$XML::Xerces::XMLUni::fgXercesValidationErrorAsFatal",  0 );
            $parser->setFeature( "$XML::Xerces::XMLUni::fgSAX2CoreValidation",            1 );
            $parser->setFeature( "$XML::Xerces::XMLUni::fgXercesDynamic",                 1 );
        };
        XML::Xerces::error($@) if $@;

        eval {
            my $is = XML::Xerces::LocalFileInputSource->new($main::FILE);
            $parser->parse($is);
        };
        XML::Xerces::error($@) if $@;

        ############
        percentDone;
        ############

        my $error_type = $XMLErrorHandler::error_type;

        if ( $error_type eq 'Warnings' || $error_type eq 'Errors' ) {

            my $message = ":-O  " . "$eml_file" . " was created - EML " . "$error_type" . " detected!";
            $lb_out->insert( "end", $message );
            $lb_out->insert( "end", "     Please see the messages above or the error log  ( " . "$cwd" . "/error.log" . " ) for details." );

        }

        else {
            my $message = ":-)  " . "$eml_file" . " was created - No EML errors detected";
            $lb_out->insert( "end", $message );
        }

        $XMLErrorHandler::error_type = 0;
    }
#################################
#  End of EML Schema Validation #
#################################

    # EML Schema option not selected

    else {
        my $message = ":-)  " . "$eml_file" . " was created";
        $lb_out->insert( "end", $message );
    }

    $lb_out->insert( "end", " " );
    $files_done = sprintf( "%.0f", $files_done );
    return $files_done;

}
