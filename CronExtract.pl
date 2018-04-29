#!/usr/bin/perl

use strict;
use File::Basename;
use Spreadsheet::WriteExcel::Big;
use Spreadsheet::XLSX::Fmt2007;


my $Cronhash;
my $Rangehash;
my $count = 0;
my $Lrange = $ARGV[0];
my $Urange = $ARGV[1];


if(defined($Lrange)){
	if ($Lrange =~ /^[a-zA-Z]+/){
		print "Lower Range Should be Numeric\n";
		exit;
	}else{
		if (length($Lrange) > 2){
			print "Range Cannot be greater then 2 digits\n";
			exit;
		}
		if($Lrange =~ /^[3,4,5,6,7,8,9]/){
				print "Lower Range cannot be greater the 22\n";
				exit;
		}elsif($Lrange =~ /^[+,\-,\.,\\,:,@,\$,\?,\,,\#,\%,\&,\*,(,),\],\[,\{,\},\=]+/){
			print "Lower Range cannot contain special characters\n";
			exit;
		}else{
			if ($Lrange =~ /^2[3,4,5,6,7,8,9]$/){
				print "Lower Range cannot be greater the 22\n";
			}
		}
	}
}else{
	print "#######################################################\n";
	print "##Usage: perl CronExtract.pl <LowerRange> <UpperRange>\n";
	#print "##Usage: perl CronExtract.pl <LowerRange> [finds jobs between at particular hour]"
	print "########################################################\n";
}


if(defined($Urange)){
	if ($Urange =~ /^[a-zA-Z]+/){
		print "Upper Range Should be Numeric\n";
		exit;
	}else{
		if (length($Urange) > 2){
			print "Range Cannot be greater then 2 digits\n";
			exit;
		}
		if ($Urange =~ /^[3,4,5,6,7,8,9]/){
				print "Upper Range cannot be greater the 23\n";
				exit;
		}elsif($Urange =~ /^[+,\-,\.,\\,:,@,\$,\?,\,,\#,\%,\&,\*,(,),\],\[,\{,\},\=]+/){
			print "Upper Range cannot contain special characters\n";
			exit;
		}else{
			if ($Urange =~ /^2[4,5,6,7,8,9]$/){
				print "Upper Range cannot be greater the 23\n";
				exit;
			}
		}
	}
}else{
	print "#######################################################\n";
	print "##Usage: perl CronExtract.pl <LowerRange> <UpperRange>\n";
	#print "##Usage: perl CronExtract.pl <LowerRange> [finds jobs between at particular hour]"
	print "########################################################\n";
}


&HashCreate;
&FindJobsBetweenRange($Lrange,$Urange);


sub HashCreate{
	
	my $user = `whoami`
	chomp($user);
	
	if(defined($Urange)){	
		open (DATA, "crontab  -l  |") ||  print "Error $? \n";
		
			while(<DATA>){
				if(!/^\#/){
				chomp;
					my ($min,$hou,$dom,$mon,$dow) = (split(/\s+/))[0,1,2,3,4];
					my ($length) = length($min) +  length($hou)  + length($dom) + length($mon) + length($dow) +5;
					my $temp = $_;
					my ($cmd) = substr($temp, $length);
					
					if (defined($cmd)){
						$Cronhash->{$user}->{$cmd}->{"min"} = $min;
						$Cronhash->{$user}->{$cmd}->{"hou"} = $hou;
						$Cronhash->{$user}->{$cmd}->{"dom"} = $dom;
						$Cronhash->{$user}->{$cmd}->{"mon"} = $mon;
						$Cronhash->{$user}->{$cmd}->{"dow"} = $dow;
					}
				}
			}
			close(DATA);
		}else{
			print("Invalid User, Kindly check the user you have logged in with");
		}
}	


sub Write2Excel{
my ( $hash ) = @_;
my $row = 0;
my $col = 0;
my $filename = "CronTab_Range.xls";
my $workbook = Spreadsheet::WriteExcel::Big->new("$filename");
my $worksheet1 = $workbook->add_worksheet("Cron");

my $format = $workbook->add_format();
$format->set_text_wrap();
my $format2 = $workbook->add_format(); # Add a format
$format2->set_bold();
$format2->set_bg_color(204,255,204);
$format2->set_align('center');

$worksheet1->write($row,$col,"User",$format2);
$col++;
$worksheet1->write($row,$col,"JobName",$format2);
$col++;
$worksheet1->write($row,$col,"Minute",$format2);
$col++;
$worksheet1->write($row,$col,"Hour",$format2);
$col++;
$worksheet1->write($row,$col,"DayOfMonth",$format2);
$col++;
$worksheet1->write($row,$col,"Month",$format2);
$col++;
$worksheet1->write($row,$col,"DayofWeek",$format2);
$col = 0;
$row = 1;

foreach my $user (sort keys %{$hash}){
		chomp($user);
		foreach my $cmd (sort keys %{$hash->{$user}}){
			$worksheet1->write($row,$col,$user);
			chomp($cmd);
			$col++;
			$worksheet1->write($row,$col,$cmd,$format);
			$col++;
			$worksheet1->write($row,$col,$hash->{$user}->{$cmd}->{"min"});
			$col++;
			$worksheet1->write($row,$col,$hash->{$user}->{$cmd}->{"hou"});
			$col++;
			$worksheet1->write($row,$col,$hash->{$user}->{$cmd}->{"dom"});
			$col++;
			$worksheet1->write($row,$col,$hash->{$user}->{$cmd}->{"mon"});
			$col++;
			$worksheet1->write($row,$col,$hash->{$user}->{$cmd}->{"dow"});
			$row++;
			$col = 0;
		}
	}
}


sub FindJobsBetweenRange{

	my ($lowerRange , $upperRange) = @_;

	foreach my $user (sort keys %{$Cronhash}){
			
			chomp($user);
			
			foreach my $range ($lowerRange .. $upperRange){
				foreach my $cmd (sort keys %{$Cronhash->{$user}}){
					
					if(exists ($Rangehash->{$user}->{$cmd})){
								next;
					}else{
									if($Cronhash->{$user}->{$cmd}->{"hou"} == $range){
											$Rangehash->{$user}->{$cmd}->{"min"} = $Cronhash->{$user}->{$cmd}->{"min"};
											$Rangehash->{$user}->{$cmd}->{"hou"} = $Cronhash->{$user}->{$cmd}->{"hou"};
											$Rangehash->{$user}->{$cmd}->{"dom"} = $Cronhash->{$user}->{$cmd}->{"dom"};
											$Rangehash->{$user}->{$cmd}->{"mon"} = $Cronhash->{$user}->{$cmd}->{"mon"};
											$Rangehash->{$user}->{$cmd}->{"dow"} = $Cronhash->{$user}->{$cmd}->{"dow"};
									}elsif($Cronhash->{$user}->{$cmd}->{"hou"} =~ /[\-]/){
										my ($first, $last) = split(/[\-]/,$Cronhash->{$user}->{$cmd}->{"hou"});
											if($range > $first && $range < $last){
												$Rangehash->{$user}->{$cmd}->{"min"} = $Cronhash->{$user}->{$cmd}->{"min"};
												$Rangehash->{$user}->{$cmd}->{"hou"} = $Cronhash->{$user}->{$cmd}->{"hou"};
												$Rangehash->{$user}->{$cmd}->{"dom"} = $Cronhash->{$user}->{$cmd}->{"dom"};
												$Rangehash->{$user}->{$cmd}->{"mon"} = $Cronhash->{$user}->{$cmd}->{"mon"};
												$Rangehash->{$user}->{$cmd}->{"dow"} = $Cronhash->{$user}->{$cmd}->{"dow"};
									}
								}elsif($Cronhash->{$user}->{$cmd}->{"hou"} =~ /[\,]/){
									my @hourRange = split(/\,/,$Cronhash->{$user}->{$cmd}->{"hou"});
										foreach my $var (@hourRange){
											if($var == $range){
												$Rangehash->{$user}->{$cmd}->{"min"} = $Cronhash->{$user}->{$cmd}->{"min"};
												$Rangehash->{$user}->{$cmd}->{"hou"} = $Cronhash->{$user}->{$cmd}->{"hou"};
												$Rangehash->{$user}->{$cmd}->{"dom"} = $Cronhash->{$user}->{$cmd}->{"dom"};
												$Rangehash->{$user}->{$cmd}->{"mon"} = $Cronhash->{$user}->{$cmd}->{"mon"};
												$Rangehash->{$user}->{$cmd}->{"dow"} = $Cronhash->{$user}->{$cmd}->{"dow"};
											}else{
												next;
											}
										}
									}
								}
							}
					}
		}
		&Write2Excel($Rangehash);
}


sub  trim {
	 my $s = shift;
	 $s =~ s/^\s+|\s+$//g;
	 return $s
}