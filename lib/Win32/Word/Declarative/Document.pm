package Win32::Word::Declarative::Document;

use warnings;
use strict;

use base qw(Class::Declarative::EventContext Class::Declarative::Node);

use File::Spec;

our $ACCEPT_EVENTS = 1;


=head1 NAME

Win32::Word::Declarative::Document - implements a document in the declarative Word framework.

=head1 VERSION

Version 0.01

=cut

our $VERSION = '0.01';


=head1 SYNOPSIS

This defines the semantics for a Word document in a declarative framework.  There are two essentially complementary modes for any
such set of semantics (and note that I'm still feeling my way here): these are I<expressive>, in which we specify a data structure that
will be written later (see L<Win32::Word::Writer>), and I<interpretive>, in which we specify, for lack of a better description, a way to
interpret an existing data structure.  The interpretive mode is what matching engines (like regexes) work with - when I specify a pattern,
I'm telling the engine how I expect to see a given set of data.

Interpretive mode is hard.  It's exciting, though, because I hope to make it an application of my concept of the active map.  That is, once
I define a matching structure for a given document, I now have a data structure that reflects that document.  If I change the data structure,
I change the document (at least potentially).  That's going to be interesting if I ever manage to get it working.

As of this writing, I'm focusing on the expressive mode, because that's today's need.  But I<let it not be forgotten> that document domains
like L<Win32::Word::Declarative::Document> will eventually contain interpretive code as well.

The document node is an event context.

There are three basic ways to use a document node.  First, it can be used with or without a file name to create a new Word document.  If it
doesn't have a filename, it'll be "Document1" or a higher number, depending on what Word assigns to it.  The second mode is to open an
existing document file.  And the third mode is to attach to a Word document already open in the user's Word session.

=head1 INTERNALS

=head2 defines()

Called by C<Class::Declarative> during import, to find out what xmlapi tags this plugin claims to implement.

=cut
sub defines { ('document'); }

=head2 build_payload, go

These functions are callbacks for the declarative framework.

=cut

sub build_payload {
   my ($self) = @_;
   
   # Set up event context for components.
   $self->event_context_init;
   $self->{app} = $self->root()->semantic_handler('ms-word')->Application; # Initialize OLE application if not already done,
                                                                           # stash for ease of use.
   $self->{constants} = $self->root()->semantic_handler('ms-word')->constants;
   $self->{callable} = 1;
   
   if ($self->parameter("active")) {  # If this node corresponds to the currently active document in Word, then we'll link to
                                      # it during the build phase.  If it's a document we're going to be creating, that's the run phase.
      $self->{payload} = $self->{app}->ActiveDocument();
      
      $self->{selection} = $self->{app}->Selection;
      my $cx = $self->event_context;
      $cx->setvalue ('app',       $self->{app});
      $cx->setvalue ('word',      $self->root()->semantic_handler('ms-word'));
      $cx->setvalue ('selection', $self->{selection});
      $cx->setvalue ('document',  $self->{payload});
   }
}

sub go {
   my ($self) = @_;
   
   $self->{app}->{Visible} = -1 if $self->parameter("visible") or $self->parameter("keepopen") or $Win32::Word::Declarative::flags->{noquit};
   
   my $file_existed = 1;
   if (not $self->parameter("active")) { # If we're not the active document, then we have to make the document.
      $self->{file} = File::Spec->rel2abs($self->label);
      if (-f $self->{file} and not $self->parameter("new")) {
         -r $self->{file} or Croak("Couldn't open file " . $self->{file} . "\n");
         $self->{payload} = $self->{app}->Documents->Open($self->{file});
      } else {
         $self->{payload} = $self->{app}->Documents->Add() or Croak ("Could not add Word document\n");
         $file_existed = 0;
      }
   }
   
   $self->{selection} = $self->{app}->Selection;
   my $cx = $self->event_context;
   $cx->setvalue ('app',       $self->{app});
   $cx->setvalue ('word',      $self->root()->semantic_handler('ms-word'));
   $cx->setvalue ('selection', $self->{selection});
   $cx->setvalue ('document',  $self->{payload});
   
   # Now process all our children.
   foreach ($self->nodes) {
      $_->go (@_);
   }

   # Finally, if we made this document and it's not marked (keepopen), then save it before leaving.
   if (not $self->parameter('active') and not $self->parameter('keepopen') and $self->{file}) {
      $self->{payload}->SaveAs($self->{file});
   }
}

=head1 EVENT CONTEXT OVERRIDES

=head2 semantics()

We return the 'ms-word' semantic handler as our core semantics.

=cut

sub semantics {
   my $self = shift;
   $self->root()->semantic_handler('ms-word');
}


=head1 DOCUMENT-SPECIFIC FUNCTIONS

=head2 selection()

Returns the current selection of the document.

=cut

sub selection { $_[0]->{selection} }



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

1; # End of Win32::Word::Declarative::Document
