<FEARANDLOATHING> Multithreading in VB </FEARANDLOATHING>

Ok, before you can even start this project you'll have to compile them
In order, open, and compile the following projects to DLLs.

1) AstarBaseClasses.vbp
2) Xtimers.vbp
3) PathFinding.vbp

Don't worry too much if you have an error like "Unable to set version
compatible component".

Then, open PathFind2.vbp and run it, making sure you have the correct
references.

You must have the Enterprise edition of Visual Basic (5 or 6) for this 
to work correctly.

This project is NOT for newbies to Visual Basic!!!!!!!!!!

Credits:

All the multithreading parts are adapted from Microsoft's multi-threaded
Coffee example, supplied with Visual Basic, and much puzzling over what
the manual has to say about it.

================

I'm attempting this project 'cos I'm sick of using DoEvents and having
applications freeze and hang because a task takes a little longer than
expected.  I also want to do more than one thing at once.

I'm also tired of seeing multithreading in VB never attempted, or if it
is, done badly using API calls.

VB is not guaranteed thread-safe (parts of it are), so I don't see why 
I should deliberately try to break it.  Instead, let's do it the way MS
recommends.

Which is not easy...sigh

This will have to be made up of 3 layers (4 if you count Xtimers):

/------------------------\
|  Main Program          |
|------------------------|
| Multi-Threaded Pieces  |
|------------------------|
| Base Classes & Types   |
\------------------------/

Why do it like this?

Erm, because the multi-threaded pieces must relate to the main program.
The "Base Classes & Types" will work sort of like a C header file and sort
of like a Windows type library.
Because of the need to pass objects between the main program & the 
multithreaded pieces, I'll need to have the base classes as a completely
seperate DLL.

Please note that VB is NOT the first computer language I've learnt.
I moved from Modula-2 (a sort of Pascal-like language where you _have_
to follow the rules of "good programming"), Unix C-Programming,
Informix-4gl, and finally Visual Basic.
I love this entry about Basic from the jargon file, and I can't resist 
quoting it here:
----
BASIC /bay'-sic/ n. 

A programming language, originally designed for Dartmouth's experimental
timesharing system in the early 1960s, which for many years was the 
leading cause of brain damage in proto-hackers. Edsger W. Dijkstra observed 
in "Selected Writings on Computing: A Personal Perspective" that "It is 
practically impossible to teach good programming style to students that 
have had prior exposure to BASIC: as potential programmers they are mentally 
mutilated beyond hope of regeneration." This is another case (like Pascal) 
of the cascading lossage that happens when a language deliberately designed 
as an educational toy gets taken too seriously. A novice can write short 
BASIC programs (on the order of 10-20 lines) very easily; writing anything 
longer (a) is very painful, and (b) encourages bad habits that will make it 
harder to use more powerful languages well. This wouldn't be so bad if 
historical accidents hadn't made BASIC so common on low-end micros in the 1980s. 
As it is, it probably ruined tens of thousands of potential wizards. 

[1995: Some languages called `BASIC' aren't quite this nasty any more, having 
acquired Pascal- and C-like procedures and control structures and shed their 
line numbers. --ESR] 

Note: the name is commonly parsed as Beginner's All-purpose Symbolic 
Instruction Code, but this is a backronym. BASIC was originally named Basic, 
simply because it was a simple and basic programming language. Because most 
programming language names were in fact acronyms, BASIC was often capitalized 
just out of habit or to be silly. No acronym for BASIC originally existed or 
was intended (as one can verify by reading texts through the early 1970s). 
Later, around the mid-1970s, people began to make up backronyms for BASIC 
because they weren't sure. Beginner's All-purpose Symbolic Instruction Code is 
the one that caught on. 
--------------------------------------------------------------------------------


Before even attempting multithreading you must

a) Disown Visual Basic.  Curse it.  Learn to despise its' intricacies
and magic way of doing things.  A quote from Bruce McKinley sums VB
up nicely:

----
While later versions of Visual Basic (4 through 6) had some 
object-oriented features, the language itself was not based on these 
features. It was based on magic. Something happened to make all those 
controls appear on forms and interact with your code. But you couldn't 
really tell what it was. And you certainly couldn't mess with it. 
That was the appeal of the language. It just worked. You could produce 
amazingly powerful applications in a short time using techniques that 
felt just right, even though they didn't make sense if you looked at 
them too closely. Now it's true that at some point you ran out of 
magic. The more complicated your application, the more likely you 
were to hit inconsistencies and incompleteness. Features added in the 
later versions of VB didn't quite fit the original plan. 
----

When you do multithreading, you are completely out of magic.  Learn to
praise Java where multithreading is so intrinsic to the language that
it becomes easy.  Kneel at the altar of C++ and swear that as soon as
you can convince your boss, you will stop using Visual Basic.  Praise
Dennis Ritchie, and sacrifice a virgin at midnight (I didn't say it 
was going to be all bad now did I)...

b) Understand threading - Read up on it - this will help you achieve
the first goal!

c) Write your code DEFENSIVELY. Error traps are essential.

d) Work out how you handle errors (I still haven't figured this one
out yet - events are probably the best way).

e) Test test test test test test test.

f) Did I mention you need to test it rigorously?

