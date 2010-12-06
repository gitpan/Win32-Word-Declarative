package Win32::Word::Declarative::Text;

use warnings;
use strict;

use base qw(Class::Declarative::Node);
use Data::Dumper;

=head1 NAME

Win32::Word::Declarative::Text - implements a text extent in the declarative Word framework.

=head1 VERSION

Version 0.01

=cut

our $VERSION = '0.01';


=head1 SYNOPSIS

The C<text> tag defines, well, text in a Word document.  It will be subclassed for paragraph and table cells at least.

The C<formatting> tag is I<also> text - but it doesn't claim its content, which is presumed to be hierarchically ordered text
elements.  This tag is thus generally used to carry formatting (hence the name) that applies to its content.

Expect at least some of this code to be factored out of Win32::Word::Declarative entirely - as soon as I need it elsewhere.
It may very well roll back into Class::Declarative.  But for now, here it is.

=head1 INTERNALS

=head2 defines()

Called by C<Class::Declarative> during import, to find out what xmlapi tags this plugin claims to implement.

=cut
sub defines { ('text', 'formatting', 'cell', 'para'); }
our %build_handlers = ( text => { node => sub { Win32::Word::Declarative::Text->new (@_) }, body => 'none' } );

=head2 build_payload

The C<build_payload> function is then called when this object is built.  All it really does is interpret its style parameters.

=cut

sub build_payload {
   my ($self) = @_;

   $self->{callable} = 1;
   $self->{style} = $self->root()->semantic_handler('ms-word')->get_style($self);
}

=head2 go

The C<go> function does the real work of creating text and applying formatting.  It is recursive to children of this node.

=cut

sub go {
   my ($self) = @_;
   
   # Our event context is the document (actually, I suppose, it will end up being the story), which contains its selection.
   # TODO: what happens if we are a documentless snippet that expects a document at run time?
   my $cx = $self->event_context;
   my $word = $self->root()->semantic_handler('ms-word');
   
   my $style_undo = $word->set_style($cx, $self->{style});
   
   # By default, we're in expressive mode and thus we add the text from our body, which is internally unparsed.
   $self->add_content($self->label ? $self->label : $self->body) unless $self->is ("formatting");
   
   # Now handle any children (note: this presumes we're a "formatting" tag or somebody is doing some macro magic;
   # "normal" text nodes won't have children).
   foreach ($self->nodes) {
      my $return = $_->go (@_);
   }

   $cx->selection->TypeParagraph if $self->is('para');

   $word->set_style($cx, $style_undo);
}

=head1 TEXT-SPECIFIC FUNCTIONS

=head2 add_content($text)

Types the text in question.  This can be overridden to create a new paragraph or table cell first, which is why it's split out like this.

=cut

sub add_content {
   my ($self, $text) = @_;
   my $cx = $self->event_context;
   $cx->selection->TypeText($text);
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

1; # End of Win32::Word::Declarative::Text
