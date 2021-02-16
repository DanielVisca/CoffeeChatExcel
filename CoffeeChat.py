import csv
import random
import json
import openpyxl
from pathlib import Path
import os
import io
import tkinter as tk
# try:
#     # for Python2
#     from Tkinter as tk   ## notice capitalized T in Tkinter 
# except ImportError:
#     # for Python3
#     from tkinter as tk 


class CoffeeChat:
    """
    db format:
    {
        people: [Person Object, ...],
        previous_matches: [Match_object]
    }
    """
    def __init__(self):
        self.upload() # creates self.people and self.previous_matches
        self.update_people()
        self.gui()
    
    def gui(self):
        window = tk.Tk()
        label = tk.Label(text="Hello, Aidan! Coffee Chat time :)")
        label.pack()
        window.mainloop()
        button = tk.Button(
            text="Generate new partners!",
            width=10,
            height=5,
            command=self.random_match()
        )
        button1 = tk.Button(
            text="Save partners to excel",
            width=10,
            height=5,
            command=self.save()
        )
        button.pack()
        button1.pack()
        
        
        # self.print_matches()
        # print(self.test())

    def save(self):
        data = {
            'people': [person for person in self.people],
            'previous_matches': [match for match in self.previous_matches]
        }
        with open('database.json','w') as f: 
            json.dump(data, f, indent=4) 
        
        self.write()
    

    def update_people(self):
        """
        Requires more extensive testing!

        If there are new people, create them and add them to the list.
        If people have been removed. remove them from the list
        """
        people = self.get_people_excel()
        # check excel against db list of people. This will find new additions
        for person in people: 
            if person not in self.people: # uncertain if 'in' will work with custom class
                print("Adding: ", person)
                self.people.append(person)
        
        # check sb list against updated excel. This will find removed people?
        for person in self.people:
            if person not in people:
                self.people.remove(person)
                print("Removed: ", person)


    def get_people_excel(self):
        xlsx_file = Path('CoffeeChat.xlsx')
        self.wb_obj = openpyxl.load_workbook(xlsx_file)
        sheet = self.wb_obj['enrolled']
        people = []
        for row in sheet.iter_rows(max_row=sheet.max_row):
            person = Person(row[0].value, row[1].value, row[2].value ,row[3].value)
            people.append(person)
        return people

    def upload(self):
        if os.path.isfile('database.json') and os.access("database.json", os.R_OK):
            print("file exists and is readable")
            with open('database.json') as f:
                db = json.load(f)
                try:
                    self.people = list(db['people'])
                    self.previous_matches = list(db['previous_matches'])
                except:
                    print("Could not read 'People' and 'Previous Matches' from database.")
                    exit()
        else:
            print ("Either file is missing or is not readable, creating file...")
            with io.open(os.path.join('database.json'), 'w') as db_file:
                db_file.write(json.dumps({}))
                self.people = []
                self.previous_matches = []
    
    def match(self, p1, p2):
        """
        Check if p1 and p2 have been matched before. If they have, return None, if they have not, create a new Match obj and
        add it to self.matches (which will be added to db)
        """
        new_match = Match(p1, p2)
        e3, e4 = new_match['emails']
        for match in self.previous_matches:
            e1, e2 = match['emails']
            
            # also set( of size 2 or less)
            # Check if both emails match
            if ((e1 == e3) or (e1 == e4)) and ((e2 == e3) or (e2 == e4)):
                return False
    
        return new_match
    

    def random_match(self):
        """
        Randomly match email addresses
        """
        match_list = []
        people_to_match = self.people.copy()
        unmatched = None

        while len(people_to_match) > 1:
            person = random.choice(people_to_match)
            # can I remove a whole dict from a list?
            people_to_match.remove(person)

            random_partner = random.choice(people_to_match)
            match = self.match(person, random_partner)

            seen_everyone = []
            tried = [random_partner]
            while not match: # already matched previously
            # maybe shift this to lower in the while loop
                if len(tried) >= len(people_to_match):
                    seen_everyone.append(person)
                    print(person['email'], " has matched with everyone already and cannot find a partner")
                    break
                random_partner = random.choice(people_to_match)
                if random_partner not in tried:
                    tried.append(random_partner)
                    match = self.match(person, random_partner)
            if match:
                print("found match")
                print(person['email'], ' ', random_partner['email'])
                # By this point they have never been matched before
                people_to_match.remove(random_partner)
                good_match = Match(person, random_partner)
                match_list.append(good_match)
                self.previous_matches.append(good_match)
        
        self.match_list = match_list

    def print_matches(self):
        print("\nThis weeks matches:\n")
        for match in self.match_list:
            print(match["emails"])
    
    def test(self):
        possible_matches = 0
        for i in range(1, len(self.people)):
            possible_matches += i
        return len(self.previous_matches) == possible_matches
    
    def write(self):
        ws1 = self.wb_obj.create_sheet("week")
        for match in self.match_list:
            people1 = match['people'][0]
            people2 = match['people'][1]
            row = (people1['email'], people1['first_name'], people1['last_name'], people1['position'], '', people2['email'], people2['first_name'], people2['last_name'], people2['position'])
            ws1.append(row)
        self.wb_obj.save("CoffeeChat.xlsx")

def Person(email, first_name, last_name, position):
    return {
        'email': email,
        'first_name': first_name,
        'last_name': last_name,
        'position': position
    }

def Match(p1,p2):
    return {
        "emails": (p1['email'], p2['email']),
        "people": (p1, p2)
    }

if __name__ == '__main__':
    cc = CoffeeChat()
