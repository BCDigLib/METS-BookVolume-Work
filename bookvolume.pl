#!C:/Perl/bin/perl -w
use strict;

use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
use IO::File;

use Cwd;


my ($sec,$min,$hour,$mday,$mon,$yr,$wday,$yday,$isdst)=localtime();
my $time=($yr+1900)."-".sprintf("%02d",$mon+1)."-".sprintf("%02d",$mday)."T".sprintf("%02d",$hour).":".sprintf("%02d",$min).":".sprintf("%02d",$sec);

my $defaultpage;

$Win32::OLE::Warn = 3; # Die on Errors.


# ::Warn = 2; throws the errors, but #
# expects that the programmer deals  #


#First, we need an excel object to work with, so if there isn't an open one, we create a new one, and we define how the object is going to exit

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
        || Win32::OLE->new('Excel.Application', 'Quit');

#For the sake of this program, we'll turn off all those pesky alert boxes, such as the SaveAs response "This file already exists", etc. using the DisplayAlerts property.

$Excel->{DisplayAlerts}=0;   

#get excel file name from ARGV

my $excelfile=shift(@ARGV);
my $dir = getcwd;
$dir=~s/\//\\/g;
print "dir is $dir\n";
$excelfile=$dir."//".$excelfile;

#opened an excel file to work with 

                                                 
 my $Book = $Excel->Workbooks->Open($excelfile);   

#Now we create a reference to a worksheet object and activate the sheet to give it focus so that actions taken on the workbook or application objects occur on this sheet unless otherwise specified.

    my $Sheet = $Book->Worksheets("Sheet1");
       $Sheet->Activate();  



#Open the output file; print xml declaration and root node
#

(my $outputfile = $excelfile) =~ s/\.xls/\.xml/;
my $fh = IO::File->new($outputfile, 'w')
	or die "unable to open output file for writing: $!";

##get structMapTopDivType

my $structMapTopDivType= shift(@ARGV);

metsOpening(); 

amdDigiProv();

##read rows and generate filegrp

my $usedRange = $Sheet->UsedRange()->{Value};


###generate filegroups

my $mimetype;
shift (@$usedRange);
shift (@$usedRange);


		$fh->print("\t<mets:fileSec ID=\"FSD1\">\n");
	if ($Sheet->Range("J2")->{Value}){
		$fh->print("\t\t<mets:fileGrp USE=\"archive image\" VERSDATE=\"".$time."\">\n");
		$mimetype="tiff";
		fileGroup($mimetype);
		$fh->print("\t\t<\/mets:fileGrp>\n");
	};

	if ($Sheet->Range("N2")->{Value}){
		$fh->print("\t\t<mets:fileGrp USE=\"reference image\" VERSDATE=\"".$time."\">\n");
        	$mimetype="j2k"; 
		fileGroup($mimetype);
		$fh->print("\t\t<\/mets:fileGrp>\n");
	};


	$fh->print("\t\t<mets:fileGrp USE=\"reference image\" VERSDATE=\"".$time."\">\n");
        $mimetype="jpeg"; 
	fileGroup($mimetype);
	$fh->print("\t\t<\/mets:fileGrp>\n");
	

	if ($Sheet->Range("M2")->{Value}){
		$fh->print("\t\t<mets:fileGrp USE=\"alto\" VERSDATE=\"".$time."\">\n");
        	$mimetype="alto"; 
		fileGroup($mimetype);
		$fh->print("\t\t<\/mets:fileGrp>\n");
	};

####adding pdf filegrp
	if ($Sheet->Range("L3")->{Value}){
		$fh->print("\t\t<mets:fileGrp USE=\"pdf\" VERSDATE=\"".$time."\">\n");
#
		$fh->print("\t\t\t<mets:file ID=\"pdf01000\" MIMETYPE=\"application/pdf\" ADMID=\"DP03\" GROUPID=\"GID1000\" SEQ=\"1000\">\n");
		$fh->print("\t\t\t\t<mets:FLocat xlink:href=\"file://streams/".$Sheet->Range("L3")->{Value}."\" LOCTYPE=\"URL\"\/>\n");
		$fh->print("\t\t\t<\/mets:file>\n");


#
		$fh->print("\t\t<\/mets:fileGrp>\n");
	};

####finished adding pdf fileGrp




$fh->print("\t</mets:fileSec>\n");

###find label for structMaps
my $label= $Sheet->Range("A1")->{Value};

##add structMap Counter
my $structMapCounter=0;

####generate mixed structMap

if ($Sheet->Range("B2")->{Value}=~ m/yes/) {

$structMapCounter=$structMapCounter+1;
$fh->print("\t<mets:structMap TYPE=\"mixed\" LABEL=\"Read Online\" ID=\"SMD".$structMapCounter."\">\n");

$fh->print("\t\t<mets:div TYPE=\"$structMapTopDivType\" LABEL=\"$label\" ORDER=\"1\" DMDID=\"DMD1\">\n");



   structMapMixed();

 $fh->print("\t\t</mets:div>\n");
 $fh->print("\t</mets:structMap>\n");

}


####generate physical structMap

$structMapCounter=$structMapCounter+1;
$fh->print("\t<mets:structMap TYPE=\"physical\" LABEL=\"Read Online\" ID=\"SMD".$structMapCounter."\">\n");

$fh->print("\t\t<mets:div TYPE=\"$structMapTopDivType\" LABEL=\"$label\" ORDER=\"1\" DMDID=\"DMD1\">\n");


   structMapPhysical();

 $fh->print("\t\t</mets:div>\n");
 $fh->print("\t</mets:structMap>\n");


####generate logical structMap

if ($Sheet->Range("D2")->{Value}=~ m/yes/) {
$structMapCounter=$structMapCounter+1;
$fh->print("\t<mets:structMap TYPE=\"logical\" LABEL=\"Read Online\" ID=\"SMD".$structMapCounter."\">\n");

$fh->print("\t\t<mets:div TYPE=\"$structMapTopDivType\" LABEL=\"$label\" ORDER=\"1\"  DMDID=\"DMD1\">\n");



   structMapLogical();

 $fh->print("\t\t</mets:div>\n");
 $fh->print("\t</mets:structMap>\n");

}



####generate pdf structMap

if ($Sheet->Range("L3")->{Value}) {
	$structMapCounter=$structMapCounter+1;
	$fh->print("\t<mets:structMap TYPE=\"pdf\" LABEL=\"PDF version\" ID=\"SMD".$structMapCounter."\">\n");
	$fh->print("\t\t<mets:div TYPE=\"$structMapTopDivType\" LABEL=\"$label\" ORDER=\"1\" DMDID=\"DMD1\">\n");

	$fh->print("\t\t\t\t<mets:fptr FILEID=\"pdf01000\"/>\n");

	 $fh->print("\t\t</mets:div>\n");
 	 $fh->print("\t</mets:structMap>\n");

}






behaviorSec();
metsClosing();

$fh->close();
#######

sub metsOpening {

my $label = $Sheet->Range("A1")->{Value};


$fh->print("<?xml version='1.0' encoding='UTF-8' ?>\n");
$fh->print("<mets:mets OBJID=\"\" LABEL=\"$label\" TYPE=\"text-monographic-whole\" xmlns:mets=\"http://www.loc.gov/METS/\"
    xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"
    xmlns:mods=\"http://www.loc.gov/mods/v3\"
    xsi:schemaLocation=\"http://www.loc.gov/METS/ http://www.loc.gov/standards/mets/mets.xsd http://www.loc.gov/mods/v3 http://www.loc.gov/standards/mods/v3/mods-3-4.xsd\">
\n");




$fh->print("<mets:metsHdr CREATEDATE=\"".$time."\">\n");
$fh->print("\t<mets:agent ROLE=\"CREATOR\" TYPE=\"ORGANIZATION\">\n");
$fh->print("\t\t<mets:name>Boston College, University Libraries, Systems Office<\/mets:name>\n");
$fh->print("\t<\/mets:agent>\n");


$fh->print("<\/mets:metsHdr>\n");
$fh->print("<!--<mets:dmdSec ID=\"DMD1\">\n");
$fh->print("\t\t<mets:mdWrap MDTYPE=\"MODS\">\n");
$fh->print("\t\t\t<mets:xmlData>\n");
$fh->print("\t\t\t<\/mets:xmlData>\n");
$fh->print("\t\t<\/mets:mdWrap>\n");
$fh->print("\t<\/mets:dmdSec>-->\n");};


#########
sub metsClosing{ $fh->print("<\/mets:mets>\n");};



###########
sub fileGroup {

my $type=shift;
my $i=1;

foreach my $row (@$usedRange) {


my ($title,$mixed,$mixed2, $logical,$ill,$label,$orderlabel,$default, $filename, $md5, $pdf, $jpeg, $alto, $j2k) = @$row;

	if ($default && ($default =~ m/yes/)) {$defaultpage = "jpg".sprintf("%05d", $i)}

	if ($type eq 'jpeg')

			{
			$fh->print("\t\t\t<mets:file ID=\"jpg".sprintf("%05d", $i)."\" MIMETYPE=\"image\/jpeg\" ADMID=\"DP02\" GROUPID=\"GID".$i."\" SEQ=\"$i\">\n");
			$fh->print("\t\t\t\t<mets:FLocat xlink:href=\"file://streams/".$filename.".jpg\" LOCTYPE=\"URL\"\/>\n"); 
			$fh->print("\t\t\t<\/mets:file>\n");
			
			$i++;}
	if ($type eq 'tiff')
			{$fh->print("\t\t\t<mets:file ID=\"tif".sprintf("%05d", $i)."\" MIMETYPE=\"image\/tiff\" ADMID=\"DP01\" GROUPID=\"GID".$i."\" SEQ=\"$i\">\n");
			$fh->print("\t\t\t\t<mets:FLocat xlink:href=\"file://streams/" . $filename . ".tif\" LOCTYPE=\"URL\"\/>\n"); 
			$fh->print("\t\t\t<\/mets:file>\n");

			$i++;}
	if ($type eq 'alto')
			{$fh->print("\t\t\t<mets:file ID=\"alto".sprintf("%05d", $i)."\" MIMETYPE=\"text\/xml\" ADMID=\"DP04\" GROUPID=\"GID".$i."\" SEQ=\"$i\">\n");

			$fh->print("\t\t\t\t<mets:FLocat xlink:href=\"file://streams/" .$filename . "_ALTO.xml\" LOCTYPE=\"URL\"\/>\n"); 
			$fh->print("\t\t\t<\/mets:file>\n");

			$i++;}

	if ($type eq 'j2k')
			{$fh->print("\t\t\t<mets:file ID=\"j2k".sprintf("%05d", $i)."\" MIMETYPE=\"image\/jp2\" ADMID=\"DP05\" GROUPID=\"GID".$i."\" SEQ=\"$i\">\n");
			$fh->print("\t\t\t\t<mets:FLocat xlink:href=\"file://streams/" . $filename . ".jp2\" LOCTYPE=\"URL\"\/>\n"); 
			$fh->print("\t\t\t<\/mets:file>\n");

			$i++;}

		

};
};

####
####

sub structMapPhysical {
my $div=1;
my $filecount=1;
my $label;
foreach my $row (@$usedRange) {
	my ($title, $mixed, $mixed2,$logical,$ill,$label,$orderlabel,$default,$filename) = @$row;

	if ($label =~ m/^\d+$/) {$label = "p. ".$label}
	if ($ill) {$label=$label ." (".$ill.")"};

	$fh->print("\t\t\t<mets:div TYPE=\"page\" LABEL=\"".$label. "\" ORDERLABEL=\"".$orderlabel."\" ORDER=\"".$div."\">\n");


		if ($Sheet->Range("J2")->{Value}) {

		$fh->print("\t\t\t\t<mets:fptr FILEID=\"tif".sprintf("%05d",$filecount)."\"/>\n")
		};
		if ($Sheet->Range("M2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"alto".sprintf("%05d",$filecount)."\"/>\n")};
		if ($Sheet->Range("N2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"j2k".sprintf("%05d",$filecount)."\"/>\n")};

		$fh->print("\t\t\t\t<mets:fptr FILEID=\"jpg".sprintf("%05d",$filecount)."\"/>\n"); 



		$fh->print("\t\t\t<\/mets:div>\n");
		$div++;
		$filecount++;

			}

}

####
####

sub structMapLogical {

my $div=1;
my $filecount=1;
my $label;
foreach my $row (@$usedRange) {
	my ($title, $mixed, $mixed2,$logical,$ill,$label,$orderlabel,$default,$filename) = @$row;
	if ($ill =~ m/yes/) {$logical=$logical." (illustration)"};	

	$fh->print("\t\t\t<mets:div TYPE=\"logical\" LABEL=\"".$logical. "\" ORDERLABEL=\"".$orderlabel."\" ORDER=\"".$div."\">\n");
		$fh->print("\t\t\t\t<mets:fptr FILEID=\"tif".sprintf("%05d",$filecount)."\"/>\n");
		if ($Sheet->Range("M2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"alto".sprintf("%05d",$filecount)."\"/>\n")};

		if ($Sheet->Range("N2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"j2k".sprintf("%05d",$filecount)."\"/>\n")};
		$fh->print("\t\t\t\t<mets:fptr FILEID=\"jpg".sprintf("%05d",$filecount)."\"/>\n"); 






		$fh->print("\t\t\t<\/mets:div>\n");
		$div++;
		$filecount++;
			}

}




######
sub structMapMixed {
my @logical_open = qw( closed closed ); #array to represent open/closed status of two logical div levels (both content)
my $mixed_div_count=1;
my $mixed2_div_count=1;
my $div_count=1;
my $filecount=1;
my $label;

foreach my $row (@$usedRange) {

	my ($title, $mixed, $mixed2,$logical,$ill,$label,$orderlabel,$default,$filename) = @$row;
	if ($label =~ m/^\d+$/) {$label = "p. ".$label}
	if ($ill) {$label=$label ." (".$ill.")"};
	
	if ($mixed && $mixed2) {	
		if ($logical_open[0] eq 'open') {$fh->print("\t\t\t\t<\/mets:div>\n")};
		if ($logical_open[1] eq 'open') {$fh->print("\t\t\t\t<\/mets:div>\n")};
		## set div status to open

		$logical_open[0] = 'open';
		$logical_open[1] = 'open';

		$div_count=1;
		$mixed2_div_count=1;
		$fh->print("\t\t\t<mets:div TYPE=\"logical\" LABEL=\"".$mixed. "\" ORDERLABEL=\"".$mixed_div_count."\" ORDER=\"".$mixed_div_count."\">\n");
		$mixed_div_count++;
		$fh->print("\t\t\t<mets:div TYPE=\"logical\" LABEL=\"".$mixed2. "\" ORDERLABEL=\"".$mixed2_div_count."\" ORDER=\"".$mixed2_div_count."\">\n");
		$mixed2_div_count++;
		$fh->print("\t\t\t<mets:div TYPE=\"page\" LABEL=\"".$label. "\" ORDERLABEL=\"".$orderlabel."\" ORDER=\"".$div_count."\">\n");
		
if ($Sheet->Range("J2")->{Value}) {

		$fh->print("\t\t\t\t<mets:fptr FILEID=\"tif".sprintf("%05d",$filecount)."\"/>\n")
		};

		if ($Sheet->Range("M2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"alto".sprintf("%05d",$filecount)."\"/>\n")};
		if ($Sheet->Range("N2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"j2k".sprintf("%05d",$filecount)."\"/>\n")};

		$fh->print("\t\t\t\t<mets:fptr FILEID=\"jpg".sprintf("%05d",$filecount)."\"/>\n"); 



		$fh->print("\t\t\t\t<\/mets:div>\n");
		$div_count++;
		$filecount++;



	}


##
	if ((! $mixed) && ($mixed2)) {


		if ($logical_open[1] eq 'open') {$fh->print("\t\t\t\t<\/mets:div>\n")};
		## set div status to open


		$logical_open[1] = 'open';
	

		$div_count=1;

		$fh->print("\t\t\t\t<mets:div TYPE=\"logical\" LABEL=\"".$mixed2. "\" ORDERLABEL=\"".$mixed2_div_count."\" ORDER=\"".$mixed2_div_count."\">\n");
		$mixed2_div_count++;
		$fh->print("\t\t\t<mets:div TYPE=\"page\" LABEL=\"".$label. "\" ORDERLABEL=\"".$orderlabel."\" ORDER=\"".$div_count."\">\n");
		if ($Sheet->Range("J2")->{Value}) {
		$fh->print("\t\t\t\t<mets:fptr FILEID=\"tif".sprintf("%05d",$filecount)."\"/>\n")
		};
		if ($Sheet->Range("M2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"alto".sprintf("%05d",$filecount)."\"/>\n")};

		if ($Sheet->Range("N2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"j2k".sprintf("%05d",$filecount)."\"/>\n")};

		$fh->print("\t\t\t\t<mets:fptr FILEID=\"jpg".sprintf("%05d",$filecount)."\"/>\n"); 

		
		$fh->print("\t\t\t\t<\/mets:div>\n");
		$div_count++;
		$filecount++;


	}
	
##
	if (($mixed) && (! $mixed2)) {

		if ($logical_open[0] eq 'open') {$fh->print("\t\t\t\t<\/mets:div>\n")};
		if ($logical_open[1] eq 'open') {$fh->print("\t\t\t\t<\/mets:div>\n")};
		## set div status to open

		$logical_open[0] = 'open';
		$logical_open[1] = 'closed';
		

		$div_count=1;
		$fh->print("\t\t\t<mets:div TYPE=\"logical\" LABEL=\"".$mixed. "\" ORDERLABEL=\"".$mixed_div_count."\" ORDER=\"".$mixed_div_count."\">\n");

		$mixed_div_count++;
		$mixed2_div_count=1;
		$fh->print("\t\t\t<mets:div TYPE=\"page\" LABEL=\"".$label. "\" ORDERLABEL=\"".$orderlabel."\" ORDER=\"".$div_count."\">\n");

		if ($Sheet->Range("J2")->{Value}) {
		$fh->print("\t\t\t\t<mets:fptr FILEID=\"tif".sprintf("%05d",$filecount)."\"/>\n")
		};

		if ($Sheet->Range("M2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"alto".sprintf("%05d",$filecount)."\"/>\n")};

		if ($Sheet->Range("N2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"j2k".sprintf("%05d",$filecount)."\"/>\n")};
		$fh->print("\t\t\t\t<mets:fptr FILEID=\"jpg".sprintf("%05d",$filecount)."\"/>\n"); 



		$fh->print("\t\t\t\t<\/mets:div>\n");
		
		$div_count++;
		$filecount++;



		

	}

	if ((! $mixed) && (! $mixed2)) {

		$fh->print("\t\t\t<mets:div TYPE=\"page\" LABEL=\"".$label. "\" ORDERLABEL=\"".$orderlabel."\" ORDER=\"".$div_count."\">\n");

		

		if($Sheet->Range("M2")->{Value}){ $fh->print("\t\t\t\t<mets:fptr FILEID=\"alto".sprintf("%05d",$filecount)."\"/>\n")};

		if ($Sheet->Range("J2")->{Value}) {
		$fh->print("\t\t\t\t<mets:fptr FILEID=\"tif".sprintf("%05d",$filecount)."\"/>\n")
		};
		
		if ($Sheet->Range("N2")->{Value}) {$fh->print("\t\t\t\t<mets:fptr FILEID=\"j2k".sprintf("%05d",$filecount)."\"/>\n")};

		$fh->print("\t\t\t\t<mets:fptr FILEID=\"jpg".sprintf("%05d",$filecount)."\"/>\n"); 

				$fh->print("\t\t\t\t<\/mets:div>\n");
		$div_count++;
		$filecount++;

		}


}
		if ($logical_open[0] eq 'open') {$fh->print("\t\t\t\t<\/mets:div>\n")};
		if ($logical_open[1] eq 'open') {$fh->print("\t\t\t\t<\/mets:div>\n")};
}

####
####

sub amdDigiProv {
		
		# hash of cells that cotain digiprove data
		my %digiProv = (  
				J2 => 'DP01' ,  #cell of tiffs used to generate derivatives
				K2 => 'DP02' ,  #cell of derivatives
				L2 => 'DP03' ,  #cell of pdf
				M2 => 'DP04' ,  #cell of alto file
				N2 => 'DP05');  #cell of uncropped masters


		$fh->print("\t<mets:amdSec ID=\"AMD1\">\n");
		
		foreach my $k (sort(keys %digiProv))
		{
			if ($Sheet->Range($k)->{Value})
			{
				$fh->print("\t\t<mets:digiprovMD ID=\"".$digiProv{$k}."\">\n");
				$fh->print("\t\t\t<mets:mdWrap MDTYPE=\"OTHER\" OTHERMDTYPE=\"local\">\n");

				$fh->print("\t\t\t\t<mets:xmlData>\n");
				$fh->print("\t\t\t\t\t<bcdigiprov>\n");
				$fh->print("\t\t\t\t\t\t<note>".$Sheet->Range($k)->{Value}."<\/note>\n");
				$fh->print("\t\t\t\t\t<\/bcdigiprov>\n");
				$fh->print("\t\t\t\t<\/mets:xmlData>\n");
				$fh->print("\t\t\t<\/mets:mdWrap>\n");
				$fh->print("\t\t<\/mets:digiprovMD>\n");
			
			}
			print"key is $k and $digiProv{$k}\n";
		};	
	



	

	


	$fh->print("\t<\/mets:amdSec>\n");

}

sub behaviorSec {

	$fh->print("\t<mets:behaviorSec>\n");
	$fh->print("\t\t<mets:behavior ID=\"BEHAVIOR1\" STRUCTID=\"$defaultpage\" BTYPE=\"DISPLAY\" LABEL=\"For each structMap, tells the system which page the book should open to as a default\">\n");
	$fh->print("\t\t\t<mets:interfaceDef LOCTYPE=\"OTHER\"\/>\n");
	$fh->print("\t\t\t<mets:mechanism LOCTYPE=\"OTHER\"\/>\n");
		
	
	$fh->print("\t\t<\/mets:behavior>\n");
	$fh->print("\t<\/mets:behaviorSec>\n");

}

=cut
usage: bookvolume.pl template.xlsx volume or manuscript
"         Remember to type volume on the command line
"         Further fill in the digiprov information



=cut