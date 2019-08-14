#!/usr/bin/perl -w

use warnings;
use strict;
use Spreadsheet::ParseXLSX;
use Getopt::Long;
use Pod::Usage;
use Term::ANSIColor;
use Config::IniFiles;
use feature qw(say);
    

sub generate{
    my ($input, $output, $debug, $fontsize, $noheader, %margins) = @_;
    my $inputxlsxfile = $input.".xlsx";
    my $outputtexfile = $output.".tex";
    open(my $fh, '>:encoding(UTF-8)', $outputtexfile);

    my $permeable = <<END;
\\documentclass[a4paper,${fontsize}pt]{extarticle}

\\usepackage[utf8]{inputenc}
\\usepackage[newdimens]{labels}

\\LabelCols=$margins{cols}                       %Number of columns of labels per page
\\LabelRows=$margins{rows}                       %Number of rows of labels per page

\\LeftPageMargin=$margins{leftPageMargin}        %These four parameters give the
\\BottomPageMargin=$margins{bottomPageMargin}    %distances from the edge of the paper.
\\RightPageMargin=$margins{rightPageMargin}      %page gutter sizes.  The outer edges of
\\TopPageMargin=$margins{topPageMargin}          %the outer labels are the specified

\\InterLabelColumn=$margins{innerLabelColumn}    %Gap between columns of labels
\\InterLabelRow=$margins{innerLabelRow}          %Gap between rows of labels

\\LeftLabelBorder=$margins{leftLabelBorder}      %These four parameters give the extra
\\RightLabelBorder=$margins{rightLabelBorder}    %space used around the text on each
\\TopLabelBorder=$margins{topLabelBorder}        %actual label.
\\BottomLabelBorder=$margins{bottomLabelBorder}  %

END

    if ($debug) {
	$permeable .= "\\LabelGridtrue            %Grid to debug\n\n";
    }

    $permeable .= "\\begin{document}\n\\begin{labels}\n";

    print $fh $permeable;

    my $parser = Spreadsheet::ParseXLSX->new;
    my $workbook = $parser->parse($inputxlsxfile);
    my $worksheet = $workbook->worksheet(0);
    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();

    if(!$noheader) { $row_min++; }

    for my $row_num ($row_min..($row_max))
    {
	for my $col_num ($col_min..($col_max)){
	    my $cell = $worksheet->get_cell($row_num,$col_num);
	    next unless $cell;
	    my $value = $cell->value();
	    print $fh $value."\n";
	}
	print $fh "\n";
    }

    print $fh "\\end{labels}\n\\end{document}\n";
    close $fh;
    system("pdflatex $outputtexfile");
    print colored(['bold green on_white'], "Done! Output file is $output.pdf.                                                                   ")."\n";
}

sub cutOffFileExtension{
    my ($filename) = @_;
    $filename =~ s/\.[^.]+$//;
    return $filename;
}

sub setMargins{
    
    my %margins;
    my ($type, $configfilename, $labelsperpage, $pagemargins, $labelmargins) = @_;

    if ($type ne "custom") {
	my $cfg = Config::IniFiles->new( -file => $configfilename );
	if (!$cfg->SectionExists($type)){die "Type not available in configuration file"}
	$margins{'cols'}	      = $cfg->val($type, 'cols');
	$margins{'rows'}	      = $cfg->val($type, 'rows');
	$margins{'leftPageMargin'}    = $cfg->val($type, 'leftPageMargin');
	$margins{'bottomPageMargin'}  = $cfg->val($type, 'bottomPageMargin');
	$margins{'rightPageMargin'}   = $cfg->val($type, 'rightPageMargin');
	$margins{'topPageMargin'}     = $cfg->val($type, 'topPageMargin');
	$margins{'innerLabelColumn'}  = $cfg->val($type, 'innerLabelColumn');
	$margins{'innerLabelRow'}     = $cfg->val($type, 'innerLabelRow');
	$margins{'leftLabelBorder'}   = $cfg->val($type, 'leftLabelBorder');
	$margins{'bottomLabelBorder'} = $cfg->val($type, 'bottomLabelBorder');
	$margins{'rightLabelBorder'}  = $cfg->val($type, 'rightLabelBorder');
	$margins{'topLabelBorder'}    = $cfg->val($type, 'topLabelBorder');
    }

    if($labelsperpage->[0]){ $margins{'cols'} = $labelsperpage->[0]; }
    if($labelsperpage->[1]){ $margins{'rows'} = $labelsperpage->[1]; }

    if($pagemargins->[0]){ $margins{'leftPageMargin'}   = $pagemargins->[0]."mm"; }
    if($pagemargins->[1]){ $margins{'bottomPageMargin'} = $pagemargins->[1]."mm"; }
    if($pagemargins->[2]){ $margins{'rightPageMargin'}  = $pagemargins->[2]."mm"; }
    if($pagemargins->[3]){ $margins{'topPageMargin'}    = $pagemargins->[3]."mm"; }
    if($pagemargins->[4]){ $margins{'innerLabelColumn'} = $pagemargins->[4]."mm"; }
    if($pagemargins->[5]){ $margins{'innerLabelRow'}    = $pagemargins->[5]."mm"; }

    if($labelmargins->[0]){ $margins{'leftLabelBorder'}   = $labelmargins->[0]."mm"; }
    if($labelmargins->[1]){ $margins{'bottomLabelBorder'} = $labelmargins->[1]."mm"; }
    if($labelmargins->[2]){ $margins{'rightLabelBorder'}  = $labelmargins->[2]."mm"; }
    if($labelmargins->[3]){ $margins{'topLabelBorder'}    = $labelmargins->[3]."mm"; }

    return %margins;
}

sub main(){

    # Loading input parameters from command line
    my $debug;
    my $noheader;
    my $configfilename = "config.ini";
    my $fontsize = 12;
    
    GetOptions(
	'debug|d'	     => \$debug,
	'input|i=s'	     => \my $input,
	'output|o=s'	     => \my $output,
	'help'		     => sub{ pod2usage(1) },
	'man'		     => sub{ pod2usage(-verbose => 2) },
	'type|t=s'	     => \my $type,
	'config=s'	     => \$configfilename,
	'labelsperpage=i{2}' => \my @labelsperpage,
	'pagemargins=f{6}'   => \my @pagemargins,
	'labelmargins=f{4}'  => \my @labelmargins,
	'noheader'	     => \$noheader,
	'fontsize=i'         => \$fontsize,
	);

    die pod2usage(1) unless $input;
    die pod2usage(1) unless $output;
    die pod2usage(1) unless $type;

    # Cut off the file extensions
    $input = cutOffFileExtension($input);
    $output = cutOffFileExtension($output);

    # Create an object to setup the margins
    my %margins = setMargins($type,$configfilename,\@labelsperpage,\@pagemargins,\@labelmargins);
    generate($input,$output,$debug,$fontsize,$noheader,%margins);
}

main();

=head1 NAME

generateLabels - generate labels from Excel(R) sheet

=head1 SYNOPSIS

generateLabels [options] --input FILE --output FILE --type TYPE

Parameters:

  --input, -i      Filename of Excel file to read from (.xlsx)
  --output, -o     Output file (.pdf)
  --type, -t       Type of labels (e.g. J8163, custom) - if custom is chosen, set up label sheet as shown below

Options:

  --debug, -d      Show grid lines on output file
  --labelsperpage  Amount of labels per page [amount_of_columns amount_of_rows]
  --pagemargins    Margins in mm of full page with labels [left bottom right top innercolumn innerrow]
  --labelmargins   Margins in mm of label [left bottom right top]
  --config         Specify configuration file if it is not the default one
  --noheader       Make also a label of the first row of the Excel spreadsheet
  --fontsize       Specify the font size in pt (default 12pt)
  --help           Print information about usage
  --man            Show man pages

=cut
