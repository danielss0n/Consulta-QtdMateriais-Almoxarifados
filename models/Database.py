import json

class Database:
    def __init__(self):
        self.db = self.get_database()

    def get_database(self):
        with open('database.json', 'r') as db:
            return json.load(db)
    
    def get_item_by_key(self, key):
        data = self.get_database()["consultas"]
        for consulta in data:
            if key in consulta:
                return consulta[key]
    
    def save_components(self, item):
         with open("database.json",'r+') as db:
            data = json.load(db)
            data["consultas"].append(item)
            db.seek(0)
            json.dump(data, db, indent = 4)