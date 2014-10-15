  ~~ weeehhh ~~
  
  #!/usr/bin/perl -w

    use strict;
    use Excel::Writer::XLSX;

    # Create a new workbook called simple.xlsx and add a worksheet
    my $workbook  = Excel::Writer::XLSX->new( 'simple.xlsx' );
    my $worksheet = $workbook->add_worksheet();

    # The general syntax is write($row, $column, $token). Note that row and
    # column are zero indexed

    # Write some text
    $worksheet->write( 0, 0, 'Hi Excel!' );


    # Write some numbers
    $worksheet->write( 2, 0, 3 );
    $worksheet->write( 3, 0, 3.00000 );
    $worksheet->write( 4, 0, 3.00001 );
    $worksheet->write( 5, 0, 3.14159 );


    # Write some formulas
    $worksheet->write( 7, 0, '=A3 + A6' );
    $worksheet->write( 8, 0, '=IF(A5>3,"Yes", "No")' );


    # Write a hyperlink
    my $hyperlink_format = $workbook->add_format(
        color     => 'blue',
        underline => 1,
    );

    $worksheet->write( 10, 0, 'http://www.perl.com/', $hyperlink_format );
