package Win32::Word::Declarative::Table;

use warnings;
use strict;

use base qw(Class::Declarative::Node);
use Data::Dumper;

=head1 NAME

Win32::Word::Declarative::Table - implements a table in the declarative Word framework.

=head1 VERSION

Version 0.01

=cut

our $VERSION = '0.01';


=head1 SYNOPSIS

The C<table> tag defines a table in a Word document.

=head1 INTERNALS

=head2 defines()

Called by C<Class::Declarative> during import, to find out what xmlapi tags this plugin claims to implement.

=cut
sub defines { ('table'); }

=head2 build_payload

The C<build_payload> function is then called when this object is built.  It really does very little at this stage.

=cut

sub build_payload {
   my ($self) = @_;

   $self->{callable} = 1;
}

=head2 go

The C<go> function does the real work of creating the table.  The table will already be built with the correct number of rows and cells
(if possible).  The number of cells in the first row will be the determining factor for now - ultimately, that may change, but there will have
to be some rational way of specifying table layout first.

=cut

sub go {
   my ($self) = @_;
   
   # Our event context is the document (actually, I suppose, it will end up being the story), which contains its selection.
   # TODO: what happens if we are a documentless snippet that expects a document at run time?
   my $cx = $self->event_context;
   my $word = $self->root()->semantic_handler('ms-word');
   my $const = $word->constants();

   # Count the number of rows and the number of cells in the first row.
   my $row_count = 0;
   my $col_count = 0;
   foreach my $r ($self->nodes) {
      $row_count += 1 if ($r->is('row'));
      if ($r->is('row') and not $col_count) {
         foreach my $c ($r->nodes) {
            $col_count += 1 if ($c->is('cell'));
         }
      }
   }
   
   # Add the table itself.  This will be more complex later; Word tables are feature-laden.
   my $table = $self->{payload} = $cx->payload->Tables->Add ({Range=>$cx->selection->range, NumRows=>$row_count, NumColumns=>$col_count});
   
   # Handle the default borders for the table.
   $word->set_border ($table, $self->parm_css ('border'));
   
   # Format the columns, if any are specified.
   my $c = 0;
   foreach my $column ($self->nodes('column')) {
      $c += 1;
      my $word_column = $table->Columns($c);
      $word_column->Select;
      $word->set_style($cx, $word->get_style($column));
      if (my $w = $column->parameter('width')) {
         my $units = 'pt';
         $units = 'in' if ($w =~ /in$/);
         $w =~ s/[ a-z]*$//i;
         $w *= 72 if $units eq 'in';
         $word_column->{PreferredWidthType} = $word->const('wdPreferredWidthPoints');
         $word_column->{PreferredWidth} = $w;
      }
   }
   
   # Now handle the children - which are, of course, the rows.
   my $r = 0;
   foreach my $row ($self->nodes('row')) {
      $r += 1;
      my $word_row = $table->Rows($r);
      $word_row->Select;
      $word->set_style($cx, $word->get_style($row));
      $word->set_border($word_row, $row->parm_css ('border'));
      my $c = 0;
      foreach my $cell ($row->nodes('cell')) {
         $c += 1;
         my $word_cell = $word_row->Cells($c);
         $word_cell->Select;
         $word->set_style($cx, $word->get_style($cell));
         $word->set_border($word_cell, $cell->parm_css('border'));
         $word_cell->Range->Select;
         $cell->go();
      }
   }
   
   # Finally, leave the insertion point after the table.
   #$cx->selection->Range} = $self->payload->{Range};
   $self->payload->Range->Select;
   $cx->selection->Collapse($cx->{constants}->{wdCollapseEnd});
}

=head1 TEXT-SPECIFIC FUNCTIONS

=head2 add_content($text)

Types the text in question.  This can be overridden to create a new paragraph or table cell first, which is why it's split out like this.

=cut

sub add_content {
   my ($self, $text) = @_;
   my $cx = $self->event_context;
   $cx->selection->TypeTable($text);
}

=head1 AUTHOR

Michael Roberts, C<< <michael at vivtek.com> >>

=head1 BUGS

Please report any bugs or feature requests to C<bug-wx-definedui at rt.cpan.org>, or through
the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=Win32-Word-Declarative>.  I will be notified, and then you'll
automatically be notified of progress on your bug as I make changes.

=head1 LICENSE AND COPYRIGHT

Copyright 2010 Michael Roberts.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.

=cut

1; # End of Win32::Word::Declarative::Table
