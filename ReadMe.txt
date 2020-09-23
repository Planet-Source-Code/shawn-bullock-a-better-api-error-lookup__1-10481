
This is just a simple program that lets you enter in an error number, and get the error description (Windows API errors).  You can also enter in a search range and it will list all the errors in that range.  Sometimes, it will want to raise an exception error, so I have some code in here that will trap them and continue.  On that note, I must admit, I don't recall which issue, but I got the source for that from a VBPJ article, therefore it's not really my source code, nor is it modified.  The rest of the source code I wrote.  Also, on that note, there are some windows errors that will not be displayed, such as number 34.  It's a disk error, but it also raises an exception when that number is queried, therefore shows as a blank.

I actually used this for my own purposes here at work.  There is another programmer here who kept pushing it past it's limits and telling me how much better I can make it, so that's how it got to be what it is today (and why it's so fast, earlier, it was painstakingly slow).  In the future, I will actually take the time to document all of the related constants and perhaps generate some VB code to deal with API errors and exception handling (try reading about Windows exception handling in some of the online MSJ articles, it's complicated).  Actually, more or less, I'll write a parser that reads into the MSDN files to retrieve them.


To lookup up only an error, just enter the error number and press enter.
To lookup a range, enter in the range number and press enter.
To Filter a search, enter in the word (single word, no expressions) and press enter
To search with a filter, enter in a filter word, and then click "List Them"

Please note, that when focus is set to a particular text field, the default button is also set so that the Enter key will yield the proper result.

Guess that's it... if there are bugs or suggestions for improvements, let me know...
leabre@hotmail.com
Shawn