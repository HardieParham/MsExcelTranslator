# Standard imports
import re
import os
import logging


# External Imports
import googletrans
import openpyxl
import openpyxl.utils
from openpyxl.utils import get_column_letter


# Local Imports
from data.settings import data, char_to_replace




class Text():
    def __init__(self, dest, src):
        #self.translator = googletrans.Translator()
        # Future TODO, setup the ability for text formatting memorization.
        # As of now, any updated text reformats text to document default
        self.alignment = ''
        self.font = ''
        self.font_size = ''
        self.content = ''
        self.dest = dest
        self.src = src


    ### 3 - create translating function
    def translate(self, text):
        try:
            for key, value in char_to_replace.items():
                new_text = text.replace(key, value)
            translation = googletrans.Translator(new_text, dest=self.dest, src=self.src)
            return translation.text
        except:
            logging.warning(f'Translation failed for {text}')





    def text_check(phrase, dic):
        # If phrase is blank, skip the check
        if phrase == None or phrase == '':
            return phrase

        else:
            # Cycle through all key,value pairs of data
            for key, value in dic.items():

                # Searching the phrase for the given key, returns the index where the key starts and ends. If no match, start_index is returned as -1
                start_index = phrase.find(key)
                end_index = start_index + len(key)

                # If a match was found, update key, value pair
                if start_index != -1:
                    print(f'MATCH! Changing {key} to {value}')

                    # If the match is at the start of the textline
                    if start_index == 0:
                        splits = phrase.split(key)
                        phrase = value + splits[1]

                    # If the match is at the end of the textline
                    elif end_index == len(phrase):
                        splits = phrase.split(key)
                        phrase = splits[0] + value

                    # If the match is somewhere in the middle
                    else:
                        splits = phrase.split(key)
                        phrase = splits[0] + value + splits[1]
                    
            # Returns back new updated phrase
            return phrase





class Workbook():
    def __init__(self, file):
        self.name = file
        self.new_name = 'ENG_' + file
        #self.new_name = Text.translate(text=self.old_name, )
        self.wb = self.load_wb()
        self.ws_titles = {}


    # Method to load the excel document to extract its information
    def load_wb(self):
        wb = openpyxl.load_workbook(f"input/{self.name}")
        self.log_change(content=f'Starting translations for {self.name}')
        return wb


    # Method to save word document after changes have been made
    def save_wb(self):
        # Hardie.wb.save(file_name)
        self.wb.save(f'output/{self.name}')


    # Method to update log.txt if any changes were made (useful if errors after script was run)
    def log_change(self, content):
        with open('data/log.txt', 'a') as f:
            f.write(f'\n {content}')


    def translate_ws_titles(self):
        for i, sheet in enumerate(self.wb.worksheets):
            print(f"{i}, {sheet.title}")
            old_title = sheet.title
            print(old_title)
            translator = Text(dest='en', src='de')
            new_title = translator.translate(text=sheet.title)
            print(new_title)
            if new_title == None:
                new_title = old_title
            if "/" in new_title:
                resulting = re.sub("/", " ", new_title)
                self.wb.worksheets[i].title = resulting
                self.titles[old_title] = resulting

            else:
                self.wb.worksheets[i].title = new_title
                self.ws_titles[old_title] = new_title
        print(self.ws_titles)


    def loop_thru_worksheet(self, sheet):
        pass


    def loop_thru_document(self):
        self.translate_ws_titles()
        for sheet in self.wb.worksheets:
            self.loop_thru_worksheet(sheet)






class App():
    def __init__(self):
        self.input_path = './input'
        self.filename_list = self.get_files()
        self.dest = data['destination language']
        self.src = data['source language']


    # Finds and registers all files in the input folder
    def get_files(self):
        filelist = []

        # registers the input directory 
        dirpath = os.listdir(self.input_path)

        # Loop through the input directory
        for file in dirpath:

            # Checks if file is valid (aka, not a subfolder)
            if os.path.isfile(os.path.join(self.input_path, file)):

                # Adds the file to the file list
                filelist.append(file)

        # return the finalized file list
        print(filelist)
        return filelist


    # Deletes log file if existing
    def delete_old_log(self):
        if os.path.exists('data/log.txt'):
            os.remove('data/log.txt')
        else:
            print('Log file not found')


    # Main scripting loop through all the files
    def main_loop(self):
        self.delete_old_log()
        
        # loop through each file in the input folder
        for file in self.filename_list:
            print(f'\n Loading {file} \n')

            # Initialize document as a word file with the script
            excel = Workbook(file)

            # Loop through the given document
            excel.loop_thru_document()

            # Save the document after the loop is completed in the output folder
            excel.save_wb()




# Start main loop when file is run
if __name__ == "__main__":
    app = App()
    app.main_loop()
    #trans = googletrans.Translator()
    #first = 'Was ist passiert?'
    #second = trans.translate(text=first, dest='en', src='de')
    #print(second.text)



# Process
# 1 - Imports
# 2 - Save new file
# 3 - create translating function
# 4 - keep dict of worksheet names
# 5 - translate worksheet names
# 6 - update dict with translations
# 7 - function to find last cell in worksheet
# 8 - translation loop through worksheet
# 9 - in text = none then pass
# 10 - if text, then translate
# 11 - if starts with =, check to see if worksheet names are in it. If yes, use dict translationj. if not then pass
# 12 - go to next worksheet
# 13 - repeat finding cell and translate
# 14 - save file again
# 15 - set settings (langauge in, languiage out)
# 16 - Start Loop