# WordToHtml

A .Net Core web api application. It shows two different ways to convert a word document to clean html:

1. Converting a word document to html by using the Mammoth library.
2. Converting a word-html document (i.e an html-file exported directly from Word) to clean and proper html by using Html agility pack.

My conclusion:
I believe the best result can be obtained by exporting the word document to a single html file (which is easy and can also be automized). And then clean the word html by using Html agility pack. It is fairly easy and you have full control.

