#!perl -T

use Test::More tests => 1;

BEGIN {
    use_ok( 'Win32::Word::Declarative' ) || print "Bail out!
";
}

diag( "Testing Win32::Word::Declarative $Win32::Word::Declarative::VERSION, Perl $], $^X" );
