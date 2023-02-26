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




# Class to handle text and translation functions
class Text():
    def __init__(self, dest, src):
        self.trans = googletrans.Translator()
        # Future TODO, setup the ability for text formatting memorization.
        # As of now, any updated text reformats text to document default
        self.alignment = ''
        self.font = ''
        self.font_size = ''
        self.content = ''
        self.dest = dest
        self.src = src


    def translate(self, text):
        try:
            # There a few characters allowed in excel that break Google Translate's API
            # So first, need to remove any instances of these, and replace with a space
            for key, value in char_to_replace.items():
                new_text = text.replace(key, value)

            # Translate new API safe string
            translation = self.trans.translate(text=new_text, dest=self.dest, src=self.src)
            return translation.text
        
        except:
            logging.warning(f'Translation failed for {text}')




class Workbook():
    def __init__(self, file):
        self.name = file
        self.new_name = 'ENG_' + file
        self.data = data
        self.wb = self.load_wb()
        self.ws_titles = {}
        self.trans = Text(dest=self.data['destination language'], src=self.data['source language'])


    # Method to load the excel document to extract its information
    def load_wb(self):
        wb = openpyxl.load_workbook(f"input/{self.name}")
        self.log_change(content=f'Starting translations for {self.name}')
        return wb


    # Method to save word document after changes have been made
    def save_wb(self):
        self.wb.save(f'output/{self.new_name}')


    # Method to update log.txt if any changes were made (useful if errors after script was run)
    def log_change(self, content):
        with open('data/log.txt', 'a') as f:
            f.write(f'\n {content}')


    # Method to find the last cell (Row and column) that a value occurs
    # NOTE: This is just the last column and last row that conatin a value. A value may not necessarily be at the row-column combination specified.
    # e.g. Last item by column is in Z1, last item by row is in B 15, This function would return Z-15, even though there is not a value in that cell.
    def get_lastcell(self):
            
            # Finds Last Entry in the last Column
            cols = (tuple(self.ws.columns))
            try:
                # Last entry of the last column of the sheet
                last = str(cols[-1][-1])

                # Seperates the last cell's Letter Column and Number Row
                # NOTE Cell starts with a '.', avoid '.' in the WS title 
                try:
                    split = last.split("'.")
                except:
                    print("WorkSheet Name Error! Change worksheet name to avoid the following character combination '.")
                
                # setting Last Cell position variable    
                position = split[1]
                row=""
                col=""

                # Seperates the Letter and Number, then assigns them to Col and Row variables respectively
                for item in position:
                    if item in ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"):
                        col += item
                    elif item == ">":
                        pass
                    else:
                        row += item

                #Converts Column string into integer
                input = col.lower()
                output = []
                for character in input:
                    number = ord(character) - 96
                    output.append(number)

                # If Column has more than 1 letter, convert to correct integer
                if len(output) > 1:
                    col_value = ((len(output) - 1) * 26) + output[-1]
                else:
                    col_value = output[0]

                #Return integers for Col and Row
                return col_value, row
            # If no cells were found, sheet is empty. Return zero coordinates.
            except:
                return 0, 0


    # Method to translate worksheet titles (the tabs at the bottom of excel)
    def translate_ws_titles(self):

        # Loop thru each sheet
        for i, sheet in enumerate(self.wb.worksheets):
            
            # Register the current title
            old_title = sheet.title

            # Translate title
            new_title = self.trans.translate(sheet.title)
            
            # If translation failed, reset title to old_title
            if new_title == None:
                new_title = old_title

            # The character '/' in a WS title crashes excel, so need to remove it    
            if "/" in new_title:
                resulting = re.sub("/", " ", new_title)

                # Update the WS title
                self.wb.worksheets[i].title = resulting

                # Update our dictionary, keeping track of the changes made
                self.ws_titles[old_title] = resulting
                self.log_change(content=f'Changed WS Title {old_title} to {new_title}')

            # If no special cases, update the WS title
            else:
                # Update the WS title
                self.wb.worksheets[i].title = new_title

                # Update our dictionary, keeping track of the changes made
                self.ws_titles[old_title] = new_title
                self.log_change(content=f'Changed WS Title {old_title} to {new_title}')


    def loop_thru_worksheet(self, lcol, lrow):
        #Loop through the rows of the worksheet
        for row in range(1,(lrow+1)):

            #Nested loop through the columns of the worksheet first
            for col in range(1,(lcol+1)):

                #Get Column text character (e.g column 'AB' = 28)
                char = get_column_letter(col)

                #Extracts the contents of the cell
                text = str(self.ws[char + str(row)].value)


                # If content starts with '=', check to see if worksheet names are in the equation. If yes, use dict translation. if not then pass
                if text[0] == "=":
                    translated_check = False
                    self.log_change(content=f'Checking if external sheet reference at {char},{row}')

                    # Loop through the stored Worksheet translations from before
                    for key, value in self.ws_titles.items():
                        if key in text:

                            # Replace found key with already translated WS title
                            translation = text.replace(key, ("'" + value + "'"))
                            self.log_change(content=f'{char}{row} was changed to {translation}')

                            # Reset the cell value with new contents
                            self.ws[char + str(row)].value = translation
                            translated_check = True

                    # If no keys were found in the equation, then pass        
                    if translated_check == False:
                        self.log_change(content=f'{char}{row}was not translated')


                #If text exists, then translate
                elif text != "None":
                    translation = self.trans.translate(text)
                    self.log_change(content=f'{char}{row} - {translation}')
                    self.ws[char + str(row)].value = translation

                # If cell is empty, then pass
                else:
                    self.log_change(content=f'{char}{row} - None')


    def loop_thru_document(self):
        # Translate all Worksheet titles first
        self.translate_ws_titles()

        # Then translate the content in each Worksheet
        for i, sheet in enumerate(self.wb.worksheets):
            self.log_change(content=f'Starting to translate: {sheet}')
            print(f'Starting to translate: {sheet}')

            # Set the currnet worksheet
            self.ws = self.wb.worksheets[i]

            # find the location of the last cell in the worksheet
            (col_value, row_value) = self.get_lastcell()

            # Translate the entire sheet
            self.loop_thru_worksheet(lcol=int(col_value), lrow=int(row_value)) 
        self.log_change(content=f'Completed translating {sheet}')      
        print(f'Completed translating {self.name}')




class App():
    def __init__(self):
        self.input_path = './input'
        self.filename_list = self.get_files()


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