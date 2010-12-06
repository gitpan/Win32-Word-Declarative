use strict;
use warnings;
use ExtUtils::MakeMaker;

WriteMakefile(
    NAME                => 'Win32::Word::Declarative',
    AUTHOR              => q{Michael Roberts <michael@vivtek.com>},
    VERSION_FROM        => 'lib/Win32/Word/Declarative.pm',
    ABSTRACT_FROM       => 'lib/Win32/Word/Declarative.pm',
    ($ExtUtils::MakeMaker::VERSION >= 6.3002
      ? ('LICENSE'=> 'perl')
      : ()),
    PL_FILES            => {},
    PREREQ_PM => {
        'Test::More' => 0,
        'Class::Declarative' => 0.07,
        'Win32::OLE' => 0,
    },
    dist                => { COMPRESS => 'gzip -9f', SUFFIX => 'gz', },
    clean               => { FILES => 'Win32-Word-Declarative-*' },
);
