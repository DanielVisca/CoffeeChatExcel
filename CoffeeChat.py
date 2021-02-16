import csv
import random
import json
import openpyxl
from openpyxl import Workbook
from pathlib import Path
import os
import io
from json import JSONEncoder


class CoffeeChat:
    """
    db format:
    {
        people: [Person Object, ...],
        previous_matches: [Match_object]
    }
    """
    def __init__(self):
        self.wb = Workbook()
        self.upload() # creates self.people and self.previous_matches
        self.update_people()
        self.random_match()
        self.save()
        self.print_matches()
        print(self.test())
        
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
        wb_obj = openpyxl.load_workbook(xlsx_file)
        sheet = wb_obj['enrolled']
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
        for match in self.previous_matches:
            e1, e2 = match['emails']
            e3, e4 = new_match['emails']

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
        if len(people_to_match) % 2 != 0:
            # Odd list:
            unmatched = random.choice(people_to_match)
            people_to_match.remove(unmatched)
            print(unmatched, " does not have a partner.")
        else:
            print("everyone is matched up this week")

        while len(people_to_match) > 1:
            person = random.choice(people_to_match)
            # can I remove a whole dict from a list?
            people_to_match.remove(person)

            random_partner = random.choice(people_to_match)
            match = self.match(person, random_partner)

            tried = [random_partner]
            while not match: # already matched previously
                if len(tried) >= len(people_to_match):
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
        ws1 = self.wb.create_sheet("matches")
        for match in self.match_list:
            people1 = match['people'][0]
            people2 = match['people'][1]
            row = (people1['email'], people1['first_name'], people1['last_name'], people1['position'], '', people2['email'], people2['first_name'], people2['last_name'], people2['position'])
            ws1.append(row)
        self.wb.save("CoffeeChat.xlsx")

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

# class Person():
#     def __init__(self, email, first_name, last_name, position):
#         self.email = email
#         self.first_name = first_name
#         self.last_name = last_name
#         self.position = position
#         self.has_matched_with = []
    
#     def __eq__(self, o: object) -> bool:
#         try:
#             return o.email == self.email
#         except:
#             return False
    
#     def toJson(self):
#         return json.dumps(self, default=lambda o: o.__dict__)
    

# class Match():
#     def __init__(self, person1, person2):
#         self.emails = set(person1.email, person2.email)
#         self.people = (person1, person2)
    
#     def __eq__(self, o: object) -> bool:
#         """
#         Two match objects are equal if the emails are the same
#         """
#         if o.emails == self.emails:
#             return True
#         else:
#             return False
    
#     def toJson(self):
#         return json.dumps(self, default=lambda o: o.__dict__)

    

cc = CoffeeChat()
# p1 = Person('danman@gmail.com', 'dan', 'man', 'software engineer')
# print(p1.__dict__)
# name, last_name , email, position

# dmg file 
