# Libraries

pillow library = for clipboard access
python-docx = for the word file task
![alt text](image.png)

customtkinter- for better and modern ui
watchdog = file and folder monitoring
pyinstaller = for .exe creation
pywin32- for live ms word working
git clone https://github.com/askaresh/blip-image-captioning-api.git - image captions


# exe Command
pyinstaller --onefile --windowed --add-data "icons;icons" --icon="icons/app_icon.ico" --hidden-import="watchdog.observers.api" app.py


# Step 3: What That Command Does
--onefile: Bundles everything into a single .exe file.

--windowed: Crucial. This hides the black console window when your app runs.

--add-data "icons;icons": This is the magic. It copies your entire icons folder (with folder.png and app_icon.ico) and bundles it inside the .exe.

--icon="icons/app_icon.ico": This sets your .exe file's icon to be your app icon.

--hidden-import="watchdog.observers.api": This is a "fix" for PyInstaller. watchdog has a hidden part that PyInstaller sometimes misses, so we're telling it to be sure to include it.

app.py: Your script!





# issue 1
Hello! I see the issue right away from your log. This is a classic "race condition" and a great find!

You can see it here

New image detected in folder: ...png (Watchdog saw the file instantly)

Adding image to ...docx... (The app tried to open it instantly)

Error: (The app crashed)

The Problem: The watchdog service is so fast that it's detecting the file before your screenshot tool has finished writing and "releasing" it. Your app then tries to grab a file that is still locked (or empty), which causes an error.

The Solution: We need to tell your app to wait one second after it detects a new file, just to give the screenshot tool a chance to finish.

I've made two simple changes to your code to fix this:

I created a new function process_new_image_file that adds the 1-second delay using app.after(1000, ...).

I updated the ImageFileHandler to call this new function.

I also changed str(e) to repr(e) in your error handler. This will give us a much more detailed error message if anything else goes wrong (it will fix the "Error: " with nothing after it).


# issue2
The Problem
When your .exe runs, it creates a temporary folder (like C:\Users\surbh\AppData\Local\Temp\_MEI12345) and unpacks your icons folder inside that folder.

But our script is still telling the app to look for the icons in the original folder (e.g., E:\...). It can't find them, so the icons are blank.

We need to make our script "smart." It needs a special function that finds the icons, no matter if it's running as a .py script or as a bundled .exe.

I've added a new function called resource_path to your code. This function will 100% solve the problem.


# APP testing 
Test 1: The "Happy Path" (Clipboard Mode)
Run python app.py. The app opens.

Click "Select .docx File" and choose your test_report.docx.

The "Clipboard" mode should be selected by default. Click "Start."

Buttons Start, Browse, and the Mode Switch should all be disabled. The Stop button should be enabled.

Take a screenshot (Win+Shift+S).

Wait 2 seconds. Now, open test_report.docx.

Result: Is your screenshot at the bottom of the document? (PASS/FAIL)

Close the doc. Take 3 more screenshots quickly.

Open the doc again.

Result: Are all 3 new screenshots there, in the correct order? (PASS/FAIL)

Click "Stop." The Start, Browse, and Mode Switch buttons should all become enabled again.

Test 2: The "Happy Path" (Folder Mode)
On the app, switch the mode to "Folder."

The "Select Folder" button should become enabled. Click it and choose your test_screenshots folder.

The label should update to "Folder: test_screenshots".

Click "Start."

Find a .png or .jpg image on your computer and drag it into the test_screenshots folder.

Wait 3 seconds. Open test_report.docx.

Result: Is the image you dragged in now in your document? (PASS/FAIL)

Click "Stop."

Test 3: The "Error Handling" Tests (Most Important!)
"No Doc File" Error:

Close and restart the app.

Click "Start" without selecting a .docx file.

Result: Does an error pop-up appear saying "Please select a .docx file first!"? (PASS/FAIL)

"No Folder" Error:

Switch to "Folder" mode.

Click "Start" without selecting a folder.

Result: Does an error pop-up appear saying "Please select a folder to monitor first!"? (PASS/FAIL)

"File Locked" Error (The Big One):

Select your test_report.docx file.

Click "Start" in "Clipboard" mode.

Now, open test_report.docx in Microsoft Word. Leave it open.

Take a screenshot (Win+Shift+S).

Result: Does an error pop-up appear saying "Please make sure the file is CLOSED..."? Does the app automatically stop monitoring? (PASS/FAIL)

Close the pop-up and close Word.

Switch to "Folder" mode, select your folder, and click "Start."

Open test_report.docx in Word again.

Drag an image into your test_screenshots folder.

Result: Does the same error pop-up appear? Does the app stop correctly? (PASS/FAIL)

Test 4: The "Final Polish" Tests
Duplicate Image Test:

Start in "Clipboard" mode. Take one screenshot.

Wait 10 seconds without copying anything else.

Result: Check your doc. The app should not have added the same image multiple times. (PASS/FAIL)

Mode Switch Test:

Start monitoring in "Clipboard" mode.

Result: Is the "Clipboard/Folder" switch disabled (so you can't change it while it's running)? (PASS/FAIL)

Click "Stop."

Result: Is the switch enabled again? (PASS/FAIL)

Window Close Test:

Start monitoring in "Folder" mode.

Click the "X" button on the app window.

Result: Does the app close without any crashes or error messages in your terminal? (This tests your on_closing function). (PASS/FAIL)



