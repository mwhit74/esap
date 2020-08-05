# esap
Excel Structural Analysis Package

This is a mostly complete, i.e. sans real testing, 2D frame analysis
written in Excel with a VBA backend. 

There may be some problems with the error checking that was
implmented, user input type checking mostly. It was one of the last
things I was working on and can't remember if I completed the testing.

It has a robust solution routine with an LU decomposition with partial
pivoting written natively in VBA. It allows for rapid solution of 
multiple load cases. The code is written in an OOP style, or as close
as you can get in VBA. 

There were plans to include many additional features but I really 
don't like writing VBA or using Excel. There are places in the code
to include shear deflection of the beam elements for 2D frames, truss
elements, and 3D beam elements.

I seriously doubt that I will be working on this in the future. If 
someone paid me a respectable amount and it would remain open 
source I would consider finishing it because it is kinda cool I 
in a "You made Excel do what?!" kind of way.   

Motivation (read at your own risk):

This package was written mostly as a tounge-in-cheek exercise due
to frustrations with the employer I had at the time and their 
insistance on using Excel to do really complex bridge analysis and
design. Operating on gigs of data in Excel sheets they would leave
running over night when Python or literally anything else would
have been done in seconds or minutes. But their argument was that
"Excel is easy and programming is complicated." and I was out to 
demonstrate that you can make Excel do things it was never supposed
to do, things that are very complicated and would be 1000 times 
easier in another platform. To show them it's fucking crippling to
use a tool not meant for the job but force it to do what you want
it to do because you are too stubborn to change to something better.
 

