KenScape Navigator 2.0 Beta Release
All rights reserved
Copyright 2003

Developer: Ken Huber

Part 1: Installation

Installation Instructions:

1.  Unzip the file, which you probably already completed.

2.  Copy the entire KenScape folder into your Program Files Directory on the root of your C drive.  Your directory structure should look like this:
	C:
	  +Program Files
	      +KenScape
	         KenScape_Navigator.exe
	         Setup.bat
	         ReadMe.txt
		+Components
	            db1.mdb
		    Instructions.html
		    msflxgrd.ocx
Nothing more nothing less.  All items are necessary for proper installation and execution.  If you are a developer and already have the msflxgrd.ocx registered elsewhere then you can skip running the Setup.bat.

3.  Finally just click on the Setup.bat file, it will take less than a sec.  Click OK and start the application.


Part 2: Features
Everything is pretty self explanatory, but I'll go over a few key features.

1.  Favorites for each different user that uses the application and saving state for each user.  Depending on your login you will have different options saved automatically for you.
2.  There are four different homepage options which are a pretty nice addition to Internet Explorer's options.  You have the option to type in an address, use your last visited page, use the current page, and the new feature I added is to have the program pick a Random page from your favorites.
3.  One of the nicest features is the fact that you can block those annoying Pop-ups from happening.  However, I realized in some cases pop-ups are necessary, ie. some retail sites use popups for their shopping carts.  To combat this, I also allow the user to turn the popups on.  However, if a popup occurs, it will open in Internet Explorer not in KenScape Navigator.  It is a flaw of windows operating systems.
4.  Another option I allow you to set is how many links you want your history to save, it varies anywhere from 10 to 1000.
5.  The easiest yet seemingly most confusing option is the option that I added to automatically add .com to the end of your URL.  The way this works is if you do not have any periods within the URL it will assume you wanted to go to a .com site since they are the most popular right now.  However, if you have a period anywhere within the URL it will not add the .com.  Here is an example using this feature: if you type in "yahoo" and hit enter it will make the url "yahoo.com" and go to the site, if you type in "www.yahoo" you will get a page not found error because it will not add the .com to the end because it found a period in your string.  If you type in "yahoo.com" it will not add a period.  It only adds .com when it does not find a period in your URL.  I find this to be very convenient.
6.  The final option that I have not discussed ties into the next topic of favorites.  It is whether or not you would like your favorites to start expanded or condensed.  When you click on the View Favorites option under the Favorites menu, you will see a grid populated with your favorites and categories.  If the category is underlined, that means it is condensed, you can double click on it and expand it to view your favorites.  Categories are displayed in a bold faced font and links are displayed in normal text.

Part 3: Organizing Favorites

1.  When adding a new favorite, the program will give you a list of categories you have already created or give you the option to choose "Other".  When choosing "Other" it will unlock a text box and allow you to input the name of the new category.
2.  KenScape will automatically fill in the Title for the page, but it will be open for editting in case you want to change the title.
3.  To go to one of your favorites you can simple View Favorites from the Favorites menu or hit Ctrl+F.
4.  From the View Favorites menu, you can delete individual favorites or you can delete entire categories.

If you have any further questions you can email me at khuber0420@aol.com or you can email me to let me know what you think of my browser.