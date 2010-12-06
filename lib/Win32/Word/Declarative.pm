package Win32::Word::Declarative;

use warnings;
use strict;

use base qw(Class::Declarative::Semantics);
#use base qw(Class::Declarative::Semantics Exporter);
#use vars qw(@EXPORT);
use Iterator::Simple;
use Class::Declarative::Util;

#@EXPORT = qw(const);

use Data::Dumper;

use Carp;

use Win32::OLE;
use Win32::OLE::Const;
use File::Temp;



=head1 NAME

Win32::Word::Declarative - Provide a declarative interface to Microsoft Word documents (in and out)!

=head1 VERSION

Version 0.01

=cut

our $VERSION = '0.01';


=head1 SYNOPSIS

The business world runs on Microsoft Office, and plain old documents in that framework are composed in Word.  MS Word has
a surprisingly sophisticated architecture, designed from the ground up to support automation - granted, it's buggy as hell and often
doesn't exactly do what you want it to do, but the need to hook it into Perl is one I encounter on a regular basis.

Problem is, there aren't many CPAN modules to do that.  So I usually just end up writing and rewriting application-specific OLE-targeted
code and punting.  This module is my attempt to start getting things into a coherent framework I can work with reusably.

You might also be interested in L<Win32::Word::Writer> if you just want to generate some Word documents using a more mature module than
this one.  W32::W::W doesn't really do everything I want to today, so I'm starting from scratch (while admittedly leaning heavily on
the techniques there).  Wish me luck.

The other perceived weakness of this module is that I'm writing it in my declarative framework (see L<Class::Declarative> for more, but
as yet incomplete, information), so there's some baggage you might not want to mess with.
I've done this because I have longer-term plans for it.  Wish me luck.

You can still use it in a less declarative manner - for which, see the tutorial below - but the canonical way to call a declarative module
is using the source filter, which makes the following a complete, grammatical Perl script:

   use Win32::Word::Declarative;

   document "mytest.doc"
      para (align=center, bold) "Test document"
      table (border=single)
         row (border-bottom=double)
            cell (italic) "March"
            cell          "some stuff"
         row
            cell          "April"
            cell          "other stuff"
            
The declarative structure itself is not very programmable (yet), so unless you're doing pretty restricted things or you're generating
scripts and running them on the fly, this generative mode is not likely to excite you yet.  There are two ways that the future will bring
more of interest: a macro system in the declarative framework, and a templating system in the document framework.  I wanted to get this out
onto CPAN now because it I<does> represent a useful set of code already, but your mileage may vary.

The interesting thing that works I<right now>, however, is still pretty interesting, and that is that you can also hook into a document
(including the active document currently open in Word) and run arbitrary Perl code with some useful abbreviations.  Here's a quick example:

   use Win32::Word::Declarative;

   document (active)
      do {
         print $^document->Name() . "\n";
         my $i = 0;
         foreach my $cell ($^word->get_list($^selection->Cells)) {
            $i += 1;
            $cell->Range->{Text} = "$i)";
         }
         printf "$i cells\n";
      }

This connects to the active document, then (1) prints its name, (2) finds all the cells in the current selection, and (3) replaces the
text of each cell with numbers in sequence plus a parenthesis.  I wrote this because I was having troubles with a particular table format
and its interaction with a different tool; this replaced the field-based numbering with explicit text, resulting in less confusion.

It makes use of magic variables in the context of the document - the "$^" is the special sigil for a magic variable in the declarative
framework.  What that means here is that C<$^document>, C<$^app>, C<$^word>, and C<$^selection> do pretty much what you expect them to do.
The C<get_list> function is a utility function that iterates down a list of things in a Word (or rather OLE) collection and returns them
all in a Perl list.  It makes C<foreach> work, preserving the Perliness of this script.

The vast, vast array of features provided by Word means that this will necessarily be a work in progress for, well, forever.
With luck it will be useful enough to keep me revisiting it, but I'm already inordinately pleased with what I've done so far.  Please
feel free to drop me a line if you find it useful - or not quite useful enough, ha.

=head1 TUTORIAL

A more detailed tutorial should probably follow, but honestly, writing a tutorial is harder than writing the module.  I'm just going to
provide a brief cookbook here and hope you can follow along at home.

=head2 Various ways of calling Win32::Word::Declarative

There are a number of different ways to invoke the module.

=head3 Running silently and creating a document

By default, C<Win32::Word::Declarative> connects to an open Word application object.  If there isn't one open, it creates a new one.  But
either way, when Perl closes, "Quit" is called on the application object it holds.  This means that if you don't already have Word open,
your script will run silently, probably creating a document, and then quit and close Word behind it.  If Word is already open, your script will
still create an invisible document window (probably - depends on Word) and quit.

So the canonical way to create a document is like this:

   use Win32::Word::Declarative;
   
   document (new) "mydocument.doc"
      text "Hello, Word!"
      
This opens a new Word document, types text into it, then saves it as "mydocument.doc" in the current working directory.  If there's already 
a document of that name, it will be overwritten (because of the "new" parameter).  This is a dandy way to write scripts to generate documents,
except, as noted above, that the declarative framework doesn't really provide programmable structure yet.  The obvious remedy for this is
to write a temporary script and run it until the declarative framework matures, or run in non-filter mode, build your object in a string,
and invoke it (on which, more below).

=head3 Opening or creating a document and leaving it open for use
      
If you want to create a document that will be in a visible Word window (perhaps you want to present something that the user can save, or not
save) you can override the quit behavior like this:

   use Win32::Word::Declarative qw(noquit);
   
   document
      text "Hello, Word!"
      
Word will now open a new document in a visible window, type that text, and leave it for your further use.  If you try to close, it will do the
normal "Do you want to save?" routine.

=head3 Running without using the declarative filter

If you want to write your script in regular Perl and call functions like normal people, you can do that.

   use Win32::Word::Declarative qw(-nofilter);
   
   my $doc = Win32::Word::Declarative->node (<<EOF);
    ! document "mydoc.doc"
    !    text "Hello, Word!"
   EOF
   
   $doc->go();

(The exclamation points at the start of the lines are just to make things a little easier to see. The reader will strip initial indentation,
including one identical character on each line.  Handy little trick I learned way back.)
   
This ends up looking a little like DBI, doesn't it?  Of course, you can simplify further and do it all in one step if you
don't need to keep the code tree around for anything:

   use Win32::Word::Declarative qw(-nofilter);
   
   Win32::Word::Declarative->node (<<EOF)->go();
    ! document "mydoc.doc"
    !    text "Hello, Word!"
   EOF
   
=head3 Running without semantically significant identation

Oh, you purist.  Of I<course> there's more than one right way to do this.  Just because I like indentation doesn't mean I'm
a full-blown Pythonista.  So you can use nested arrayrefs to do the work
of the line indentation parser, thus permitting you to indent things any way you want, or of course to generate this
structure by code.  To do this, each tag line (the first line
in a node) is a string, followed in its arrayref by either another string (if you are supplying a text body) or by an arrayref of the
node's children.

The line parser will then still interpret the line you supply to build the final node.

   use Win32::Word::Declarative qw(-nofilter);
   
   my $doc = Win32::Word::Declarative->node ([
      ['document "mydoc.doc"', [
         ['text "Hello, Word!"']]
      ]
   ]);
   
   ## Or alternatively:
   my $doc = Win32::Word::Declarative->node ([['document "mydoc.doc"', [['text "Hello, Word!"']]]]);
   
   # But I find that hard to read, which is why I came up with a significantly
   # indenting structural pseudocode notation in the first place.
   
   $doc->go();
   
You can mix and match string and arrayref loading; if a text node is found anywhere in your arrayref structure, the normal parser will be
invoked on it.

=head2 Writing Perl code against Word documents

This module doesn't actually help you a lot with Perl OLE code against a Word object.  Making that work involves a lot of reading through
Word's VBA documentation and hoping you've translated it into Perl correctly, and it's hit-or-miss in many cases.  There are some special
variables defined for Perl code during the code munging process; these are C<$^word>, which corresponds to the Application object,
C<$^sel>, the current selection, and C<$^document>, the document your enclosing C<document> tag refers to.

The C<$^word> object provides a couple of helper functions that make life with OLE easier in Perl; those will be shown below.

=head3 Word constants

Word does have a lot of constants, doesn't it?  One of the helper functions provided by C<$^document> is C<const>, which looks up the numeric
values of those constants starting from a string representation of the documented name.  C<const> also has some tables of equivalent
(but shorter) names that you can use from your own programs.


=head3 Mixing code and text (setting up a selection before typing)

You can mix code and text; this is useful in an existing document if you want to leave the selection in some convenient state before letting
the structural nodes type things for you.  It's also useful if you want to use Perl to do something dynamic in the middle of typing
the boilerplate text from the declarative structure around it.

Let's look at the first case - let's use Perl to position the selection somewhere interesting, then let the text specification type there.

   use Win32::Word::Declarative;
   
   document "my_existing.doc"
      do {
         
         # Todo: finish example, sigh.

=head3 Mixing code and text (doing things during typing)

The way the document tag does its work is that it simply asks each of its children to do I<their> work.  If one of the children is a
C<do> (i.e. Perl code) then it will simply get executed in sequence.  So you can do whatever you want during execution of the document
building phase just by adding some code in the middle.  Here's an example of a template expression.  It's pretty clunky, I agree,
but it does show one way you can build a document based on information you have from somewhere else.

  Example here - later

=head2 Document structure and components

As you probably guessed, you can do somewhat more in terms of creating documents than just typing text.  However, it should be said that
Word is I<extremely> feature-rich, and so there's a lot I haven't implemented yet.  I've done basic formatted text, and tables, which
should actually be enough for many purposes (like the invoice writer I used as an initial problem).  So that leaves out sections,
headers and footers, text boxes, all the drawing commands, fields, styles (!), and honestly, a lot of the functionality of tables, too, like
shading.

You can, of course, call the OLE functions yourself in Perl for other types of document structure.  And eventually, as my needs evolve,
I'll fill in more functionality.  Again: if there's anything you'd like to see specifically, drop me a line.

=head3 The "text" tag vs. the "para" tag

The basic component of a Word document is text, of course.  But there are a couple of different tags that can represent text, and there are
two different ways for text to appear in a text node.

In the declarative framework, every node has a label (a short string in quotes on the same line as the tag) and a body (lines of text below it).
Some tags treat their body text as child nodes, and some treat it as raw text.  The C<text> tag treats it as raw text, and will type anything
you put into it, line breaks and all.

   document
      text
         This is a longer text.
         It will be spread over a couple of lines.
         
         You can use a blank line, too.
         
If the text tag has a label, it will ignore its body.

   document
      text "This text will appear, with no line break"
         Anything you type here will be ignored entirely.
         
A text node can have formatting:

   document
      text "This is normal text.  This is "
      text (italic) "italicized"
      text ".  And this is "
      text (bold) "bold text"
      text "."
      
A C<para> tag is also a text node, but it treats its body text (if any) as children, so they get parsed as nodes themselves.  It also types
an explicit paragraph marker after its text.  If there's no text, the C<para> tag just turns into a paragraph marker.

   document
      text "This is text on one line."
      para
      text "This is text on another."
      
That is I<exactly> equivalent to:

   document
      text
         This is text on one line.
         This is text on another.
         
However, to do paragraph formatting, you I<must> put the text in a paragraph tag:

   document
      para (align=center) "My heading"
      text "This is text under the heading"
      
This is because formatting on a tag is undone before you leave the tag - meaning that centered alignment is undone before you leave
the paragraph, unless you explicitly set it on a paragraph.  You're free to say C<text (align=center) "Centered text"> - but it won't center
the text.  Or rather, it won't center it long enough for you to see it centered.

The text in the paragraph tag can also be in subordinate text tags, though:

   document
     para (align=center)
        text "This is a heading "
        text (italic) "with style!"
        
There's one more text-type tag to be noted here, the C<formatting> tag, which doesn't type a paragraph marker, but still treats its body 
as children:

   document
      formatting (italic)
         text "This is an italic sentence "
         text (bold) "with some bold italic"
         text " text in it."

Clear?  Of course it is.

=head3 Tables

It's truly amazing how complicated table layout is, once you really start looking at it in detail.  Word has a I<boatload> of functionality
having to do with table layout, the surface of which I have only begun to scratch.  I can do basic tables without merged cells, but I have a
lot of control of borders working already.

The key to understanding tables is that the C<cell> tag contains text.  The C<row> tag contains cells and can be formatted (like specifying
that the whole row will be bold).  The C<column> tag does I<not> contain text; it's just there to format colums in terms of text formatting
and any specific width you want to assign.  You can leave the C<column> tags out entirely and your table will work fine.

  Example here - later

=head1 INTERNALS

You can probably ignore this stuff; it's more oriented towards the code to define the module.

=head2 import, tag, start

The start action for Word is to create a Word document.  That's kind of my default start action for document-related domains (see PDF).

=head2 Application()

The domain itself contains the machinery for working with Word (i.e. the OLE object for Word itself).  The Word Application object is the
central OLE object that represents Word itself.  C<Application> returns that object; if it's not defined, it initializes a new OLE object
and returns that.

=cut

our $flags = {};
sub import {
   my($type, @arguments) = @_;
   foreach (@arguments) {
      $flags->{$_} = 1;
   }
   
   my $caller = caller();   # Because caller() acts *weird* in list context!  Perl is so funky.
   if ($caller->can('yes_i_am_declarative')) {
      $type->scan_plugins ($caller, __FILE__);
   } else {
      if (@arguments and $arguments[0] eq '-nofilter') {
         eval "use Class::Declarative qw(-nofilter $type);";
      } else {
         eval "use Class::Declarative qw($type);";
      }
   }
}
sub tag { 'ms-word' }
our $App;
our $constants;
sub Application {
   return $App if defined $App and $App;
   $App = Win32::OLE->GetActiveObject('Word.Application');
   if (not defined $App) {
      $App = Win32::OLE->new('Word.Application', $flags->{noquit} ? undef : 'Quit')
             or die("Couldn't get an Application object");
   }
   Win32::OLE->Option(Warn => 3);
   $constants = Win32::OLE::Const->Load($App) or die("Couldn't load Word constants");
   
   return $App;
}
sub constants { $constants }
sub start {
   # Find first document node and tell it to run its action.
   my ($self) = @_;
   my $main = $self->{root}->find('document');
   $main->go() if defined $main;
}

=head1 UTILITY FUNCTIONS FOR DEALING WITH WORD

=head2 get_list ($OLElist)

Lists of objects in the OLE format are clunky; we have to retrieve the list->Count and then iterate - and even worse,
OLE assumes an index base of 1, not 0.  Ugh.  So here's a useful function to do that for you.
Instead of writing the following in an action:

   my @cells = ();
   for (my $i = 0; $i < $^selection->Cells->Count; $i++) {
      push @cells, $^selection->Cells->Item($i+1);
   }

(which is also fine if you like that) you can write:

   my @cells = $sem->get_list ($^selection->Cells)
   
It is extremely likely that this sort of thing will roll out into a more general OLE semantics.
   
=cut

sub get_list {
   my ($self, $list) = @_;
   my @returns;
   for (my $i = 0; $i < $list->Count; $i++) {
      push @returns, $list->Item($i+1);
   }
   @returns;
}

=head2 get_iterator ($OLElist), get_lazy ($OLElist)

Same as C<get_list> except that it returns a lazy iterating list.  This is useful if the collection you're working with is, say, the words
in the document (of which there are probably many).

=cut

sub get_iterator {
   my ($self, $list) = @_;
   my $count = $list->Count;
   my $i = 0;
   sub {
      return if $i >= $count;
      $i += 1;
      $list->Item($i);
   }
}

sub get_lazy { lazyiter (get_iterator (@_)) }

=head2 get_style ($node)

Given a node (generally "$self", of course), find out any style specifications it has in its parameters.  Any unspecified style parameters
are left C<undef>.  Returns a hashref.

=cut

sub get_style {
   my ($self, $node) = @_;
   
   my $output = {};
   $output->{bold}   = -1 if $node->parameter ('bold')       || $node->parameter ('b');
   $output->{bold}   =  0 if $node->parameter ('not bold')   || $node->parameter ('b-');
   $output->{italic} = -1 if $node->parameter ('italic')     || $node->parameter ('italics')     || $node->parameter ('i');
   $output->{italic} =  0 if $node->parameter ('not italic') || $node->parameter ('not italics') || $node->parameter ('i-');
   $output->{font}   = $node->parameter ('font')  if $node->parameter ('font');
   $output->{size}   = $node->parameter ('size')  if $node->parameter ('size');
   $output->{align}  = $self->const($node->parameter ('align'), 'para-align') if $node->parameter ('align');
   
   return $output;
}

=head2 set_style ($cx, $style)

Given a Word context and a style hash as gotten by C<get_style>, sets the style as indicated, returning an undo hash for restoration of the
current style.

=cut

sub set_style {
   my ($self, $cx, $style) = @_;
   
   
   my $style_undo = {};
   
   # Handle formatting.
   foreach my $f (keys %$style) {
      if ($f eq 'bold') {
         $style_undo->{bold} = $cx->selection->{Font}->{Bold};
         $cx->selection->{Font}->{Bold} = $style->{bold};
      } elsif ($f eq 'italic') {
         $style_undo->{italic} = $cx->selection->{Font}->{Italic};
         $cx->selection->{Font}->{Italic} = $style->{italic};
      } elsif ($f eq 'font') {    # TODO: maybe something a little smarter here?
         $style_undo->{font} = $cx->selection->{Font}->{Name};
         $cx->selection->{Font}->{Name} = $style->{font};
      } elsif ($f eq 'size') {
         $style_undo->{size} = $cx->selection->{Font}->{Size};
         $cx->selection->{Font}->{Size} = $style->{size};
      } elsif ($f eq 'align') {
         $style_undo->{align} = $cx->selection->{ParagraphFormat}->{Alignment};
         $cx->selection->{ParagraphFormat}->{Alignment} = $style->{align};
      }
   }

   return $style_undo;
}

=head2 const (name, type)

Word uses a I<lot> of constants.  Really a lot of constants, of the general form of wd<Type><Name>, e.g. wdLineStyleDashDotStroked.
They're exposed through the OLE constant structure, and by default the C<const> function will retrieve them by name. 
But it also provides an easy way to define some easier names, while still permitting the use of the
standard ones.  If an alias is used, the "type" must be provided as a namespace.

=cut

our $aliases = {
   'linestyle' => {
      'single' => 'wdLineStyleSingle',
      'double' => 'wdLineStyleDouble',
      'none'   => 'wdLineStyleNone',
   },
   'linewidth' => {
      '0.5' => 'wdLineWidth050pt',
   },
   'color' => {
      'auto' => 'wdColorAutomatic',
   },
   'border' => {
      'left' => 'wdBorderLeft',
      'right' => 'wdBorderRight',
      'top', => 'wdBorderTop',
      'bottom', => 'wdBorderBottom',
      'horizontal', => 'wdBorderHorizontal',
      'vertical', => 'wdBorderVertical',
   },
   'para-align' => {
      'left' => 'wdAlignParagraphLeft',
      'right' => 'wdAlignParagraphRight',
      'center' => 'wdAlignParagraphCenter',
      'justify' => 'wdAlignParagraphJustify',
   }
};

sub const {
   my ($word, $name, $type) = @_;
   return unless defined $constants;
   if (defined $type) {
      if (defined $aliases->{$type}) {
         if (defined $aliases->{$type}->{$name}) {
            $name = $aliases->{$type}->{$name};
         }
      }
   }
   croak ("No Word constant $name defined") unless defined $constants->{$name};
   return $constants->{$name};
}

=head2 set_border (object, border_spec)

Word borders are pretty complex, but they're handled the same for tables, text extents, rows, and so on.  Given such an object and a border
specification (CSS-style hierarchical hashref), this function sets everything up.

=cut

sub set_border {
   my ($word, $object, $border) = @_;
   return unless $border;
   $border = { all => $border } unless ref $border;    # A string value for border is treated as a universal alias for the border parts.
   if ($border->{all}) {   # TODO: this kind of alias logic might be worth preserving elsewhere.
      $border->{outside}    = $border->{all};
      $border->{vertical}   = $border->{all};
      $border->{horizontal} = $border->{all};
   }
   if ($border->{outside}) {
      $border->{left}   = $border->{outside};
      $border->{right}  = $border->{outside};
      $border->{top}    = $border->{outside};
      $border->{bottom} = $border->{outside};
   }

   foreach ('left', 'right', 'top', 'bottom', 'horizontal', 'vertical') {
      if ($border->{$_}) {
         my $b = $object->Borders($word->const($_, 'border'));
         my $s = $border->{$_};
         $s = { 'style' => $s } unless ref $s;  # Default simple-value is the style (single, etc.)
         $s->{style} = 'single' unless $s->{style};
         $s->{color} = 'auto' unless $s->{color};
         $s->{width} = '0.5' unless $s->{width};

         $b->{LineStyle} = $word->const($s->{style}, 'linestyle');
         if ($word->const($s->{style}, 'linestyle')) { # i.e. style is not "none" - Word can be so damned picky.
            $b->{LineWidth} = $word->const($s->{width}, 'linewidth');
            $b->{Color}     = $word->const($s->{color}, 'color');
         }
      }
   }
}


=head1 AUTHOR

Michael Roberts, C<< <michael at vivtek.com> >>

=head1 BUGS

Please report any bugs or feature requests to C<bug-win32-word-declarative at rt.cpan.org>, or through
the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=Win32-Word-Declarative>.  I will be notified, and then you'll
automatically be notified of progress on your bug as I make changes.




=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc Win32::Word::Declarative


You can also look for information at:

=over 4

=item * RT: CPAN's request tracker

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=Win32-Word-Declarative>

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/Win32-Word-Declarative>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/Win32-Word-Declarative>

=item * Search CPAN

L<http://search.cpan.org/dist/Win32-Word-Declarative/>

=back


=head1 ACKNOWLEDGEMENTS


=head1 LICENSE AND COPYRIGHT

Copyright 2010 Michael Roberts.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.


=cut

1; # End of Win32::Word::Declarative
