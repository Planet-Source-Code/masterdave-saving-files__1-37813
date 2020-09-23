<div align="center">

## Saving Files


</div>

### Description

FOR BEGGINERS - This article explains clearly about how to save files in VB 6.0, it is SUPER EASY to fallow and explains everything very well.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[MasterDave](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/masterdave.md)
**Level**          |Beginner
**User Rating**    |1.8 (18 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/masterdave-saving-files__1-37813/archive/master.zip)





### Source Code

```
Start by making a new project, Standard EXE. Make one Textbox and one Command Button<p>
----------------------------------------------<p>
Now for the code...
First make a variable called "F" as integer. <p>
<p>
Open the code window for the command button and we will review this code
Type in the code for the command button’s click event (The click event is the default event that loads when you view an objects code for the first time, so really you don't have to change it) <p>
<p>
F = FreeFile<p>
Open App.Path & "\X.txt" for Output as #F <p>
<p>
'*NOTE* The code above is only the first 2 lines<p>
<p>
"Open" - When you type this it simply means you are going to open some file to put something in the file, get something from the file, or do whatever you need to do with it. Just remember you need to open it before you can do something with it :) <p>
<p>
"App" - App just supplies information about the program that you are making. For example, one of the things it supplies is the location for your program that you will be making. <p>
<p>
"Path" - Path simply means that the string that you will write after it will be a Path or Location to a file on your computer. <p>
 <p>
"\X.txt" - This is the location of the file that you will want to open. If the file does not already exist it will create it for you. Since we typed “\X.txt” without some drive like "C:" that means the file will be created in the path/location of your program. <p>
 <p>
"For Output" - This means that you will open the file for output, output means that it will erase all the current information inside the file and replace it with whatever you want to replace it with. There is also "Append" which will continue from an existing file and not erase the data inside it. If the file does not exist, it will be created for you. <p>
<p>
"#F" - #F is just a number to refer to the file.
Would you want to always type the path of the file every time you wanted to refer to it? That’s why you can just say #F. You <p>
<p>
'*NOTE* The first line just opens the file to process with it. (File Processing) <p>
------------------------------------------<p>
3rd line of code (still in the same code window)
Type the fallowing line underneath<p>
<p>
Open App.Path & "\X.txt" for Output as #F<p>
Print #F, Text1.Text<p>
<p>
'*NOTE* this is only the 2nd line of code. <p>
<p>
Now lets review the code for this line
"Print" - This is telling your program to print something. You can also put "Write" instead of Print, but if you put "Write" you will see Quotes (" ") around your text inside the file. I suggest you put "Print".<p>
<p>
"#F," - Remember how I said this is a reference to the file? Well that’s how the "Print" statement will know what to print to. If you just put "Print Text1.Text" it will print the text on your form instead of the file. Don't forget you need the "," so you can add what you want it to print after that. <p>
 <p>
"Text1.Text" - This means that the program will print a line of whatever is currently in the text of the textbox "Text1" to the file. <p>
<p>
'*NOTE* So the 3rd line of code will just print something inside the file. <p>
'-------------------------------------------------<p>
Now the 4th line, the easiest one
Underneath the first 3 lines write: <p>
<p>
"Close #F" - This will just close the file that is currently in "#F". What was in "#F"? "\X.txt" is in #F remember? So it will just close the file. Why do we need to close the file you ask? We need to close the file so other programs can open it when they need to. Like the WordPad program might need to open it so you can view the text inside it. Also you might want to overwrite the file from your program, if you don't close it you cannot overwrite it again. Another thing, if you don't close the file, it will take up space in your system memory and will run your computer slower than it should be running. <p>
<p>
'*NOTE* The 4th line just closes the file that you opened in the first line. <p>
<p>
So all together you will have 4 lines of code, which will be: <p>
'-------------------------------------------<p>
F = FreeFile<p>
Open App.Path & "\X.txt" For Output as #F<p>
Print #F, Text1.Text<p>
Close #F <p>
'------------------------------------------- <p>
<p>
If there is anything wrong with my code please leave feedback telling me what was wrong.
If you liked it, remember you can vote for me if you want to. <p>
<p>
Please try not to insult me when leaving feedback :)
```

