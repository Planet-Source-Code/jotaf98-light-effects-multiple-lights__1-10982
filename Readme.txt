	       _______________________________
	      (				      )
	      ( DRAWLIGHT III FUNCTION README )
	      (_______________________________)
		      _____|_____|_____
		     ( Multiple Lights )
		      ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
			______________
		       ( INTRODUCTION )
			¯¯¯¯¯¯¯¯¯¯¯¯¯¯
 "DrawLight" is a function that I wrote because I wanted to
add lightning effects to my game, if it was possible. Guess
what, it is -- and here's the proof! :)

 It draws a light in any picture, but it's faster than most
ways to achive this (because I used the circle's core
equation instead of the Circle function). Also, you can
choose the color that will be used in the effect.

 It should teach you something about using equations with
loops, messing with the RGB values of each pixel to get some
nice effects, and how basic algorythms for graphics work.

 The only difference between the first version and this one
(not counting the "painting" interface, just to demonstrate
how fast this function is) is that the multiple lights
problem is fixed now (in the first 2 versions it produced
some weird colors when 2 or more lights were on top of each
other). So now it's ready to be used in a game :p

 Really thanks to Simon Price for the function that extracts
the RGBs without producing the weird colors ;)

 Oh, I almost forgot: I "migrated" to SetPixelV instead of
SetPixel because people say it's faster...
 _______
( NOTE: ) For anyone that doesn't know, I've been on holydays
 ¯¯¯¯¯¯¯ this last month. That explains my "inactivity" for
	 this long.
			 _______
		        ( BONUS )
  ___			 ¯¯¯¯¯¯¯
 (_1_) The light is drawn from the inside to the outside, like
if it was growing. So maybe there's no need for using a
backbuffer ;)
  ___
 (_2_) Included is a code snippet that explains how to draw a
circle (a sub called "DrawACircle"). Might be useful for any
other algorythms you make.
  ___
 (_3_) Another snippet is included -- when you exit the program
and it asks you "Visit http://www.planet-source-code.com
now?", click "Yes". Your browser will take you to that page,
no matter what it is! All you need is a .url file (Internet
Link), and the line you'll find in the "Form_Unload" event!
  ___
 (_4_) Now, you can draw multiple lights on top of each other,
so it could be useful in a game or in a painting program
(like this demo).

  ___
 (_5_) You can click "Predefined Colors" to get one of the ten
coolest colors I found! Try using different colors to get the
maximum out of this effect!

  ___
 (_6_) If you're still not convinced of how cool it would be for
a painting program, have a look at "Bird.gif". I renamed the
background image so the .exe wouldn't find it (making a black
background) and drew that myself. Please note that I didn't
spend even 5 minutes on it, so it's not that well drawn ;)
			______________
		       ( INSTRUCTIONS )
			¯¯¯¯¯¯¯¯¯¯¯¯¯¯
 Simply add the .bas file to your project to use it.
 All you need to do in order to activate the effect is to
call this function.
			   _______
			  ( USAGE )
			   ¯¯¯¯¯¯¯
 DrawLight(Target As Long, X As Single, Y As Single, RedB
As Byte, GreenB As Byte, BlueB As Byte, Radius As Integer,
NumberOfSteps As Single)
			 ___________
			( ARGUMENTS )
 ___			 ¯¯¯¯¯¯¯¯¯¯¯
(___) Target:
 This is the hDC where the image will be drawn.
(You'll find this as a property of forms and picture
boxes.
 Just know that it refers to the image used in these
controls.)
 ___
(___) X:
 The X coordinate of the place where the light will be
drawn (referring to the center).
 ___
(___) Y:
 The Y coordinate of the place where the light will be
drawn (referring to the center).
 ___
(___) RedB
 This is the brightness applied to the Red value. Be
aware that high numbers will make the light totally
white!
 ___
(___) GreenB
 The brightness applied to the Green value.
 ___
(___) BlueB
 The brightness applied to the Blue value.
 ___
(___) Radius:
 The radius of the light.
 ___
(___) NumberOfSteps
 The light is drawn in various steps (circles that
apply different brightnesses to the background). The
higher the number of steps, the smoother and better the
light will appear, but too many steps will make the
process slower. Usually the radius divided by 2 is the
best bet.
			  _________
			 ( EXAMPLE )
			  ¯¯¯¯¯¯¯¯¯
 Check the sample project included for a working demo.
 But the .exe runs much faster than Visual Basic's Run
Mode!

		  ________________________
		 (			  )
		 ( Made by Jotaf98        )
		 ( [João F. S. Henriques] )
		 (________________________)
		   ____|____________|___
		  ( jotaf98@hotmail.com )
		   ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯