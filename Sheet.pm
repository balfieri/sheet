# Copyright (c) 2014-2019 Robert A. Alfieri
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
#
# High-Level Spreadsheet Manipulation (.csv only for the moment)
#
# $self->{row_vals} is an array of rows.
# Each row is an array of values.  
# Row0 contains the labels.
# Each value is an array of subvalues.  The reason for the extra level here is 
# to be able to merge multiple sheets and to group subvals for the same label.
#
# For example, after merging two sheets, $self->{row_vals} looks like:
#
# [ 
#   [ [subval0_0, subval0_1], [subval1_0, subval1_1], ... ],    # row 0 (labels)
#   [ [subval0_0, subval0_1], [subval1_0, subval1_1], ... ],    # row 1
#   ...
# ]
#
# Note: a sheet with only one set of subvals will still have the extra level of array hierarchy.
#
package Sheet;

use strict;
use warnings FATAL => 'all';
use diagnostics;

require Spreadsheet::WriteExcel;                # open source package for writing .xlsx files
use Spreadsheet::WriteExcel::Utility;           # import some useful functions

my $debug = 0;
my $debug_chart = 0;

#----------------------------------------------------
# Common error function.
#----------------------------------------------------
sub fatal ($)
{
    my $msg = shift;

    die "ERROR: $msg\n";
}

#----------------------------------------------------
# Create an empty sheet.
#----------------------------------------------------
sub new
{
    my $pkg  = shift;
    my $name = shift;

    bless my $self = { name      => $name, 
                       row_vals  => [] };  
    return $self;
}

#----------------------------------------------------
# Read a sheet from a file (assuming one subval per label for now).
#----------------------------------------------------
sub read
{
    my $pkg  = shift;
    my $name = shift;
    my $file = shift;

    my $self = $pkg->new( $name );

    open( C, $file ) or fatal "could not open $file for reading";
    my $line_num = 0;
    my $line;
    while( $line = <C> )
    {
        chomp $line;
        $line =~ s/\r//g;
        $debug and print "$line\n";

        $line_num++;
        my @vals = ();
        my $in_str = 0;
        my $val = "";
        while( $line ne "" ) 
        {
            $line =~ s/^(.)//;
            my $ch = $1;
            if ( $ch eq '"' ) {
                $in_str = !$in_str;
            }

            if ( ($ch eq "," && !$in_str) || $line eq "" ) {
                $ch ne "," and $val .= $ch;
                if ( $val =~ s/^"([\d\.\,]+)"$/$1/ ) {
                    $val =~ s/,//g;
                }
                $debug and print "    $val\n";
                push @vals, [$val];     # always add the extra level of hierarchy 
                $val = "";
                $line =~ s/^\s*//;
            } else {
                $val .= $ch;
                $line eq "" and die "line should not be empty at this point";
            }
        }

        $line_num == 1 || @vals == @{$self->{row_vals}->[0]}
            or fatal "Number of values on line $line_num does not match number of column names";
        push @{$self->{row_vals}}, [@vals];     
    }
    close( C );

    return $self;
}

#----------------------------------------------------
# Write out a sheet to a file.
#----------------------------------------------------
sub write
{
    my $self = shift;
    my $file = shift;

    open( C, ">$file" ) or fatal "could not open $file for writing";
    my $row = 0;
    for my $vals ( @{$self->{row_vals}} )
    {
        my $row_str = "";
        my $col = 0;
        for my $subvals ( @{$vals} )
        {
            my $subvals_str = join( ",", @{$subvals} );
            $col != 0 and $row_str .= ",";
            $row_str .= $subvals_str;
            $col++;
        }
        print C $row_str . "\n";
        $row++;
    }
    close( C );
}

#----------------------------------------------------
# Write out a sheet to a file in pivot format.
#----------------------------------------------------
sub write_pivot
{
    my $self        = shift;
    my $file        = shift;
    my $id_label    = shift; 
    my $hdr_row_cnt = shift || 1;
    my $subcol_inc  = shift || 1;
    my $subcol_start= shift || 0;
    my $subcol_pad  = shift || 1;

    my $id_col = $self->label_col( $id_label );

    open( C, ">$file" ) or fatal "could not open $file for writing";
    my $row = 0;
    for my $row ( $hdr_row_cnt .. @{$self->{row_vals}}-1 )
    {
        my $vals = $self->{row_vals}->[$row];
        my $row_str = "";
        my $col = 0;
        for my $subvals ( @{$vals} )
        {
            if ( $col == $id_col ) {
                $col++;
                next;
            }

            my $subcol = 0;
            for( my $subcol = $subcol_start; $subcol < (@{$subvals}-$subcol_pad); $subcol += $subcol_inc )
            {
                my $subval = $subvals->[$subcol];
                my $id = $vals->[$id_col]->[$subcol];
                my $str = $id . "," . $subval;
                for my $hrow ( 0 .. $hdr_row_cnt-1 )
                {
                    my $hdr = $self->{row_vals}->[$hrow]->[$col]->[$subcol];
                    $str .= "," . $hdr;
                }
                print C $str . "\n";
            }
            $col++;
        }
        $row++;
    }
    close( C );
}

#----------------------------------------------------
# Create an Excel workbook to be associated with this.
# It will be written automatically as changes are made.
#----------------------------------------------------
sub excel_create
{
    my $self = shift;
    my $file = shift;
    my $no_write = shift || 0;

    #----------------------------------------------------
    # Use an external Perl module.
    #----------------------------------------------------
    defined $self->{workbook} and fatal "Excel workbook already created";
    $self->{workbook}  = Spreadsheet::WriteExcel->new( $file );
    $self->{worksheet} = $self->{workbook}->add_worksheet();

    #----------------------------------------------------
    # Flatten out the rows.
    # And transpose row data into column data.
    #----------------------------------------------------
    my $rows  = $self->{row_vals};
    my $fcols = [];
    my $r = 0;
    for my $row ( @{$rows} ) 
    {
        my $i = 0;
        for my $s_vals ( @{$row} ) 
        {
            for my $s_val ( @{$s_vals} ) 
            {
                $r == 0 and push @{$fcols}, [];
                my $fcol = $fcols->[$i];
                defined $fcol or fatal "row $r has more columns than previous rows";

                push @{$fcol}, $s_val;
                $i++;
            }
        }
        $i == @{$fcols} or fatal "row $r has fewer columns than previous rows";
        $r++;
    }

    #----------------------------------------------------
    # Write the data into the worksheet.
    #----------------------------------------------------
    $self->{worksheet_cols} = $fcols;
    !$no_write and $self->{worksheet}->write( "A1", $fcols );
}

#----------------------------------------------------
# Merge two sheets where there's a match
# in a certain column with a given label.
#----------------------------------------------------
sub merge
{
    my $self      = shift;
    my $label     = shift;
    my $other     = shift;

    my $col = $self->label_col( $label );

    my $s_subvals_cnt = @{$self->{row_vals}->[0]->[0]};   # assume same everywhere
    my $o_subvals_cnt = @{$other->{row_vals}->[0]->[0]};  # assume same everywhere
    my $m_subvals_cnt = $s_subvals_cnt + $o_subvals_cnt;

    my $o_unmatched = [];

    my $row = 0;
    for my $o_vals ( @{$other->{row_vals}} )
    {
        if ( $row == 0 ) {
            #----------------------------------------------------
            # column names should match for now => duplicate them
            #----------------------------------------------------
            my $s_vals = $self->{row_vals}->[0];
            @{$o_vals} == @{$s_vals} or fatal "number of columns does not match for other row 0";
            my $col = 0;
            for my $o_subvals ( @{$o_vals} ) 
            {
                my $s_subvals = $s_vals->[$col];
                push @{$s_subvals}, @{$o_subvals};  # going on blind faith
                $col++;
            }

        } else {
            #----------------------------------------------------
            # Try to find a matching row in $self.
            #----------------------------------------------------
            my $found = 0;
            for my $s_vals ( @{$self->{row_vals}} )
            {
                if ( $s_vals->[$col]->[0] eq $o_vals->[$col]->[0] ) {
                    my $col = 0;
                    for my $o_subvals ( @{$o_vals} ) 
                    {
                        my $s_subvals = $s_vals->[$col];
                        push @{$s_subvals}, @{$o_subvals};  # going on blind faith
                        @{$s_subvals} == $m_subvals_cnt or fatal "too many subvals at row $row col $col";
                        $col++;
                    }
                    $found = 1;
                    last;
                }
            }

            #----------------------------------------------------
            # If none found, then will need to add this unmatched row later.
            #----------------------------------------------------
            !$found and push @{$o_unmatched}, $o_vals;
        }

        $row++;
    }

    #----------------------------------------------------
    # First go through and find all self rows that did not match.
    # Append missing subvalues.
    #----------------------------------------------------
    for my $s_vals ( @{$self->{row_vals}} )
    {
        if ( @{$s_vals->[0]} < $m_subvals_cnt ) {
            for my $s_subvals ( @{$s_vals} ) 
            {
                while( @{$s_subvals} < $m_subvals_cnt ) 
                {
                    push @{$s_subvals}, "";
                }
            }
        }
    }

    #----------------------------------------------------
    # Next, go through all the other unmatched rows and 
    # do the same kind of thing, except prepend missing subvalues.
    #----------------------------------------------------
    for my $o_vals ( @{$o_unmatched} )
    {
        for my $o_subvals ( @{$o_vals} ) 
        {
            while( @{$o_subvals} < $m_subvals_cnt ) 
            {
                unshift @{$o_subvals}, "";
            }
        }
    }
    push @{$self->{row_vals}}, @{$o_unmatched};  # add at the bottom (unsorted)
}

#----------------------------------------------------
# Average columns, adding a new row at the top.
#----------------------------------------------------
sub average
{
    my $self        = shift;
    my $hdr_row_cnt = shift || 1;
    my $subcol_inc  = shift || 1;
    my $subcol_start= shift || 0;
    my $subcol_pad  = shift || 1;

    #----------------------------------------------------
    # Make a copy.
    # Go through all columns.
    #----------------------------------------------------
    my $averaged = $self->slice();
    my $col_cnt = int( @{$self->{row_vals}->[0]} );
    my $row_cnt = int( @{$self->{row_vals}} );
    my $subcol_cnt = int( @{$self->{row_vals}->[$hdr_row_cnt]->[0]} );
    my $averages = [];
    my $one_minus_averages = [];
    for my $col ( 0 .. $col_cnt - 1 )
    {
        my $sub_count   = [];  for my $i ( 0 .. $subcol_cnt-1 ) { $sub_count->[$i] = 0 }
        my $sub_sum     = [];  for my $i ( 0 .. $subcol_cnt-1 ) { $sub_sum->[$i]   = 0 }
        my $sub_average = [];
        my $sub_one_minus_average = [];

        for my $row ( $hdr_row_cnt .. $row_cnt-1 )
        {
            my $subvals = $self->{row_vals}->[$row]->[$col];
            for( my $subcol = $subcol_start; $subcol < ($subcol_cnt-$subcol_pad); $subcol += $subcol_inc )
            {
                my $subval = $subvals->[$subcol];
                defined $subval && ($subval =~ /^\d+(\.\d+)?$/) && $subval > 0 or next;

                $sub_count->[$subcol] += 1;
                $sub_sum->[$subcol]   += $subval;
            }
        }

        for my $subcol ( 0 .. $subcol_cnt-1 )
        {
            if ( $subcol < $subcol_start || $subcol >= ($subcol_cnt-$subcol_pad) ) {
                $sub_average->[$subcol] = "";
                $sub_one_minus_average->[$subcol] = "";
            } elsif ( $sub_count->[$subcol] == 0 ) {
                $sub_average->[$subcol] = 0.0;
                $sub_one_minus_average->[$subcol] = 1.0;
            } else {
                $sub_average->[$subcol] = $sub_sum->[$subcol] / $sub_count->[$subcol];
                $sub_one_minus_average->[$subcol] = 1.0 - $sub_average->[$subcol];
            }
        }

        push @{$averages}, $sub_average;
        push @{$one_minus_averages}, $sub_one_minus_average;
    }

    unshift @{$averaged->{row_vals}}, $averages, $one_minus_averages;
    return $averaged;
}

#----------------------------------------------------
# Compare subvals within a sheet that has been merged.
# We currently always compare to the first subval.
#----------------------------------------------------
sub compare
{
    my $self = shift;
    my $add_blank_col = shift || 1;

    #----------------------------------------------------
    # We do this for each set of subvals.
    #----------------------------------------------------
    my $row = 0;
    my $new_row_vals = [];
    for my $vals ( @{$self->{row_vals}} )
    {
        my $new_vals = [];
        for my $subvals ( @{$vals} ) 
        {
            my $new_subvals = [];
            for my $i ( 0 .. @{$subvals}-1 )
            {
                if ( $row == 0 ) {
                    push @{$new_subvals}, $subvals->[$i], "";
                } else {
                    my $base = $subvals->[0];
                    my $this = $subvals->[$i];
                    $base =~ s/,//g;
                    $this =~ s/,//g;
                    ($base =~ /^\s*(\d*\.)?\d+\s*$/) or $base = 0;
                    ($this =~ /^\s*(\d*\.)?\d+\s*$/) or $this = 0;
                    my $ratio = ($base != 0) ? ($this / $base) : 0;
                    $ratio = sprintf( "%0.2f", $ratio );
                    #print "base=$base this=$this ratio=$ratio\n";
                    push @{$new_subvals}, $subvals->[$i], $ratio;
                }
            }
            $add_blank_col and push @{$new_subvals}, "";
            push @{$new_vals}, $new_subvals;
        }
        push @{$new_row_vals}, $new_vals;
        $row++;
    }

    $self->{row_vals} = $new_row_vals;
}

#----------------------------------------------------
# Redo comparisons done previously
#----------------------------------------------------
sub redo_compares
{
    my $self        = shift;
    my $base_col    = shift;
    my $base_subcol = shift;
    my $row_first   = shift || 0;
    my $col_first   = shift || 0;
    my $kind        = shift || "ratio";

    #----------------------------------------------------
    # We do this in place.
    #----------------------------------------------------
    my $row = 0;
    for my $vals ( @{$self->{row_vals}} )
    {
        if ( $row < $row_first ) {
            $row++;
            next;
        }

        my $base = $vals->[$base_col]->[$base_subcol];
        $base =~ s/,//g;
        ($base =~ /^\s*(\d*\.)?\d+\s*$/) or $base = 0;

        my $col = 0;
        for my $subvals ( @{$vals} ) 
        {
            if ( $col < $col_first ) {
                $col++;
                next;
            }

            my $cnt = int( @{$subvals} );
            for( my $i = 0; $i < $cnt-1; $i += 2 )
            {
                my $this = $subvals->[$i];
                $this =~ s/,//g;
                ($this =~ /^\s*(\d*\.)?\d+\s*$/) or $this = 0;
                my $ratio = ($base != 0) ? ($this / $base) : 0;
                if ( $kind eq "1-ratio" ) {
                    $ratio = 1 - $ratio;
                } else {
                    $kind eq "ratio" or die "ERROR: unknown redo kind=$kind\n";
                }
                $ratio = sprintf( "%0.2f", $ratio );
                #print "row=$row col=$col subcol=$i subcol_cnt=$cnt ratio=$ratio\n";
                $subvals->[$i+1] = $ratio;
            }
            $col++;
        }
        $row++;
    }
}

#----------------------------------------------------
# Return the number of columns (vals) in each row (assuming all have the same).
#----------------------------------------------------
sub col_cnt
{
    my $self = shift;

    return int( @{$self->{row_vals}->[0]} );
}

#----------------------------------------------------
# Return the column for a given label.
# We assume that row0 has the labels.
#----------------------------------------------------
sub label_col
{
    my $self = shift;
    my $label = shift;

    my $col = 0;
    for my $vals ( @{$self->{row_vals}->[0]} )
    {
        $vals->[0] eq $label and return $col;
        $col++;
    }
    fatal "could not find label '$label' in row 0, found these labels:\n", join( " ", @{$self->{row_vals}->[0]} ) . "\n";
}

#----------------------------------------------------
# Insert a new row of values before the given row.
#----------------------------------------------------
sub row_insert_before
{
    my $self     = shift;
    my $before   = shift;
    my $new_vals = shift;

    my $new_row_vals = [];
    my $row = 0;
    for my $vals ( @{$self->{row_vals}} )
    {
        if ( $row == $before ) {
            push @{$new_row_vals}, $new_vals;
        }
        push @{$new_row_vals}, $vals;
        $row++;
    }

    $self->{row_vals} = $new_row_vals;
}

#----------------------------------------------------
# Return a slice of a sheet.
#----------------------------------------------------
sub slice
{
    my $self         = shift;
    my $row_first    = shift || 0;
    my $row_last     = shift || 0x7fffffff;
    my $col_first    = shift || 0;
    my $col_last     = shift || 0x7fffffff;
    my $subcol_first = shift || 0;
    my $subcol_last  = shift || 0x7fffffff;

    my $new_row_vals = [];
    my $row = 0;
    for my $vals ( @{$self->{row_vals}} )
    {
        if ( $row >= $row_first && $row <= $row_last ) {
            my $new_vals = [];
            my $col = 0;
            for my $subvals ( @{$vals} ) 
            {
                if ( $col >= $col_first && $col <= $col_last ) {
                    my $new_subvals = [];
                    my $subcol = 0;
                    for my $subval ( @{$subvals} )
                    {
                        if ( $subcol >= $subcol_first && $subcol <= $subcol_last ) {
                            push @{$new_subvals}, $subval;
                        }
                        $subcol++;
                    }
                    push @{$new_vals}, $new_subvals;
                }
                $col++;
            }
            push @{$new_row_vals}, $new_vals;
        }
        $row++;
    }

    my $sheet = Sheet->new( $self->{name} );
    $sheet->{row_vals} = $new_row_vals;
    return $sheet;
}

#----------------------------------------------------
# Return a slice of a sheet according to a list of labels.
#----------------------------------------------------
sub slice_by_labels
{
    my $self   = shift;
    my $labels = shift;

    my $label_cols = [];
    for my $label ( @{$labels} )
    {
        push @{$label_cols}, $self->label_col( $label );
    }

    my $new_row_vals = [];
    for my $vals ( @{$self->{row_vals}} )
    {
        my $new_vals = [];
        my $col = 0;
        for my $subvals ( @{$vals} ) 
        {
            my $keep_col = 0;
            for my $label_col ( @{$label_cols} )
            {
                $label_col == $col and $keep_col = 1;
            }

            if ( $keep_col ) {
                push @{$new_vals}, [ @{$subvals} ];
            }
            $col++;
        }
        push @{$new_row_vals}, $new_vals;
    }

    my $sheet = Sheet->new( $self->{name} );
    $sheet->{row_vals} = $new_row_vals;
    return $sheet;
}

#----------------------------------------------------
# Convert ratios to some other measurement.
#----------------------------------------------------
sub ratio_transform
{
    my $self      = shift;
    my $measure   = shift;
    my $row_first = shift || 1;
    my $row_last  = shift || 0x7fffffff;

    my $row = 0;
    for my $vals ( @{$self->{row_vals}} )
    {
        if ( $row >= $row_first && $row <= $row_last ) {
            for my $subvals ( @{$vals} ) 
            {
                my $col = 0;
                for( my $i = 1; $i < @{$subvals}; $i += 2 )
                {
                    if ( $measure eq "pct_reduction" ) {
                        if ( $subvals->[$i] eq "" || $subvals->[$i] == 0 ) {
                            $subvals->[$i] = 0.0;
                        } else {
                            $subvals->[$i] = 1.0 - $subvals->[$i];
                        }
                    } else {
                        fatal "unknown ratio transformation measure '$measure'";
                    }
                    $col++;
                }
            }
        }
        $row++;
    }
}

#----------------------------------------------------
# Sort by [label, subcol] in given direction.
#----------------------------------------------------
sub sort_by
{
    my $self   = shift;
    my $label  = shift;
    my $dir    = shift || "ascending";
    my $row_first = shift || 1;
    my $subcol = shift || 0;

    my $col = $self->label_col( $label );

    #----------------------------------------------------
    # We need to save the first unshorted rows of labels and sublabels.
    # Those are not sorted.
    #----------------------------------------------------
    my @row_vals = @{$self->{row_vals}}; 
    my @saved = ();
    for my $i ( 0 .. $row_first-1 ) {
        push @saved, shift @row_vals;
    }

    my @new_row_vals;
    if ( $dir eq "ascending" ) {
        @new_row_vals = sort { defined $a->[$col]->[$subcol] && $a->[$col]->[$subcol] ne "" or $a->[$col]->[$subcol] = 0;
                               defined $b->[$col]->[$subcol] && $b->[$col]->[$subcol] ne "" or $b->[$col]->[$subcol] = 0;
                               $a->[$col]->[$subcol] <=> $b->[$col]->[$subcol] } @row_vals;
    } else {
        $dir eq "descending" or fatal "bad sort_by direction '$dir'";
        @new_row_vals = sort { defined $a->[$col]->[$subcol] && $a->[$col]->[$subcol] ne "" or $a->[$col]->[$subcol] = 0;
                               defined $b->[$col]->[$subcol] && $b->[$col]->[$subcol] ne "" or $b->[$col]->[$subcol] = 0;
                               $b->[$col]->[$subcol] <=> $a->[$col]->[$subcol] } @row_vals;
    }

    my $sheet = Sheet->new( $self->{name} );
    $sheet->{row_vals} = [@saved, @new_row_vals];
    return $sheet;
}

#----------------------------------------------------
# Similar to sort, except discards rows that match a certain rule
#----------------------------------------------------
sub discard_by
{
    my $self    = shift;
    my $cmp     = shift;
    my $label   = shift;
    my $row_first = shift || 1;
    my $subcol  = shift || 0;

    my $col = $self->label_col( $label );

    #----------------------------------------------------
    # We need to save the first unsorted rows of labels and sublabels.
    # Those are not sorted.
    #----------------------------------------------------
    my @row_vals = @{$self->{row_vals}}; 
    my @saved = ();
    for my $i ( 0 .. $row_first-1 ) {
        push @saved, shift @row_vals;
    }

    my @new_row_vals = ();
    for my $vals ( @row_vals )
    {
        my $val  = $vals->[$col]->[$subcol];
        my $keep = eval( $cmp ) == 0;
        $keep and push @new_row_vals, $vals;
    }

    my $sheet = Sheet->new( $self->{name} );
    $sheet->{row_vals} = [@saved, @new_row_vals];
    return $sheet;
}

#----------------------------------------------------
# Write a chart to a separate Excel spreadsheet file
#
# TODO: range limits for outliers
#----------------------------------------------------
sub write_chart
{
    my $self           = shift;
    my $file           = shift;
    my $g              = shift;
    my $col_first      = shift;
    my $col_last       = shift;
    my $sort_label     = shift;
    my $sort_row_first = shift || 1;
    my $sort_subcol    = shift || 0;
    my $subsubcol_cnt  = shift || 1;

    $debug_chart and print "col_first=$col_first col_last=$col_last sort_label='$sort_label' sort_row_first=$sort_row_first sort_subcol=$sort_subcol\n";

    my $x_axis = $g->{x_axis};
    my $y_axis = $g->{y_axis};

    #----------------------------------------------------
    # Extract only the cols we need.
    #----------------------------------------------------
    defined $y_axis->{sheet_labels} or $y_axis->{sheet_labels} = [$y_axis->{sheet_label}];
    my $sheet = $self->slice_by_labels( [$x_axis->{sheet_label}, @{$y_axis->{sheet_labels}}] );
    #$sheet->write( $file . ".csv" );

    #----------------------------------------------------
    # Assume all rows have the same number of cols
    # and that all cols have the same number of subcols. 
    #----------------------------------------------------
    my $subcol_cnt       = int( @{$sheet->{row_vals}->[0]->[0]} ) / $subsubcol_cnt;
    $debug_chart and print "subcol_cnt=$subcol_cnt\n";
    my $only_sort_subcol = $g->{y_axis}->{only_sort_subcol} // 0;
    my $subcol_first     = $only_sort_subcol ? $sort_subcol : ($subcol_cnt == 1) ? 0 : $col_first;
    my $subcol_last      = $only_sort_subcol ? $sort_subcol : ($subcol_cnt-1);
    my $sort_subcol2     = $sort_subcol * $subsubcol_cnt;
    $debug_chart and print "subcol_first=$subcol_first subcol_last=$subcol_last\n";

    #----------------------------------------------------
    # If we're going by pct_reduction, then subtract the ratios from 1.0.
    # The ratio is the 2nd value for each config.
    #----------------------------------------------------
    if ( $y_axis->{measure} ne "raw" ) {
        $sheet->ratio_transform( $y_axis->{measure}, $sort_row_first );
        $sort_subcol2++;
    }

    #----------------------------------------------------
    # Sort
    #----------------------------------------------------
    my $sort_axis  = $g->{sort_axis} // "y_axis";
    $sort_label = $sort_label // $g->{$sort_axis}->{sheet_labels}->[0];
    my $sort_dir   = $g->{sort_dir} // "ascending";

    #$debug_chart and $sheet->write( $file . ".PRESORTED.csv" );
    my $sorted = $sheet->sort_by( $sort_label, $sort_dir, $sort_row_first, $sort_subcol2 );
    #$debug_chart and $sorted->write( $file . ".SORTED.csv" );

    #----------------------------------------------------
    # Discard bogus rows.
    #----------------------------------------------------
    my $test = ($y_axis->{measure} eq "pct_reduction") ? "\$val >= 0.999" : "\$val <= 0.0";
    $sorted = $sorted->discard_by( $test, $sort_label, $sort_row_first, $sort_subcol2 );
    #$debug_chart and $sorted->write( $file . ".DISCARDED.csv" );

    #----------------------------------------------------
    # The sheet's data is all set.
    # Create an Excel spreadsheet.
    #----------------------------------------------------
    $sorted->excel_create( $file, 1 );

    #----------------------------------------------------
    # Add a chart to the sheet.
    #----------------------------------------------------
    my $chart = $sorted->{workbook}->add_chart( type => $g->{kind} );
    my $title = $g->{title};
    my $x_label = $g->{x_axis}->{axis_label} // $g->{x_axis}->{sheet_label};
    my $y_label = $g->{y_axis}->{axis_label} // $g->{y_axis}->{sheet_label};
    defined $title and $chart->set_title ( name => $title );
    $chart->set_x_axis( name => $x_label );
    $chart->set_y_axis( name => $y_label );
    $chart->set_chartarea( color => "silver" );
    $chart->set_plotarea( color => "silver" );

    #----------------------------------------------------
    # Populate the chart's series.
    # There's a separate series for each [col, subcol].
    #----------------------------------------------------
    for my $col ( $col_first .. $col_last )
    {
        my $reversed = $g->{y_axis}->{reversed} // 0;
        my $col2 = $reversed ? ($col_last - $col) : $col;

        for my $subcol ( $subcol_first .. $subcol_last ) 
        {
            my $subcol2 = $subcol * $subsubcol_cnt;   # due to extra column per value 

            #----------------------------------------------------
            # X axis values are always column 0, subsubcol 0.
            #----------------------------------------------------
            my $x1 = xl_rowcol_to_cell( $sort_row_first, 0 + $subcol2 );
            my $x2 = xl_rowcol_to_cell( int( @{$sorted->{row_vals}} - 1), 0 + $subcol2 );

            #----------------------------------------------------
            # Y axis values are based on $col2 above, then subcol, then last subsubcol.
            #----------------------------------------------------
            my $xl_col = $col2*$subcol_cnt*$subsubcol_cnt - 2;
            $xl_col   += $subcol*$subsubcol_cnt;
            $y_axis->{measure} ne "raw" and $xl_col += $subsubcol_cnt - 1;

            my $y1 = xl_rowcol_to_cell( $sort_row_first, $xl_col );
            my $y2 = xl_rowcol_to_cell( int( @{$sorted->{row_vals}} - 1 ), $xl_col );

            my $index1  = $col - $col_first;
            my $label   = $g->{y_axis}->{chart_labels}->[$index1] // $g->{y_axis}->{sheet_labels}->[$index1];
            if ( defined $g->{y_axis}->{chart_sublabels} ) {
                my $index2   = $subcol2/$subsubcol_cnt - 1;
                $label .= "_" . $g->{y_axis}->{chart_sublabels}->[$index2];
            }
            defined $label && $label ne "" or die "ERROR: blank label\n";
            $debug_chart and print "col=$col col2=$col2 subcol=$subcol subcol2=$subcol2 index=$index1 $label\n";

            my $allowed = 1;
            if ( defined $g->{y_axis}->{only_labels} ) {
                $allowed = 0;
                for my $lab ( @{$g->{y_axis}->{only_labels}} )
                {
                    if ( $label eq $lab ) {
                        $allowed = 1;
                        last;
                    }
                }
            }

            !$allowed and next;

            $chart->add_series( name        => $label,
                                categories  => "=Sheet1!${x1}:${x2}",
                                values      => "=Sheet1!${y1}:${y2}" );
        }
    }

    #----------------------------------------------------
    # Write out the modified worksheet.
    #----------------------------------------------------
    $sorted->{worksheet}->write( "A1", $sorted->{worksheet_cols} );
}

1;
