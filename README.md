# MsExcelTranslator
Python script to translate excel documents through Google Translate's API

step 1:
Update 'data/settings' to set source language and destination language

step 2:
Create venv from requirements.txt

step 3:
Run main.py to update documents according to the settings.

step 4:
Enjoy the 8 hours you saved not having to do this manually



NOTES:
Many special characters break Google Translate's API functions. Edit the char_to_replace dictionary in data/settings.py if any additional characters break it as well.