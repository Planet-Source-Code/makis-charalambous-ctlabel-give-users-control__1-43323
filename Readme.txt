This control allows your users to change the label captions of your 
fields at run time. I think this is a unique aproach to customized legends.
A lot of my customers sometimes dont agree with the captions I use 
on certain forms (nitty gritty:) so I sat down one afternoon and 
cooked this activex. Is not very commented but the code is very simple
and easy to follow. This is a quicky control, and if you like it then
I will try and improve it.

It uses a simple access database called msCustom.mdb. Of course you
can change the name and add more features to it.

To use the code the only thing you need to do is to open up a connection
when you program starts and pass the recordset to the labels at the load
event of your forms. You must also provide unique id numbers to the ctlabel.id
property. These numbers is the key to the stored strings in the table. If you use
the same id on more labels then when a user changes one label then all the labels
with the same id will be changed.

When you hold down the control and double click on a ctlabel a small window
appears (cform) where you can type the new caption you want and also change
the bold, italic or underline properties of the label. If you want to revert
back to the default caption then simply delete the text after you double click
on the ctlabel and press enter. The default caption (design caption) will appear.

The control stores ONLY the captions the users change, NOT all the captions 
in your program. 

To make it easy to test the software I did not compile it to an activex but left
the control as *.ctl in the example program. To make it activex simply delete the
example.frm, goto to project properties and change the project type from Standard Exe
to ActiveX Control and compile it to an OCX. Have fun.

If you have any questions or problems then please contact me. I hope that this
small control is as useful to you as it is to me at info@mastersoft.com.cy

Makis Charalambous
