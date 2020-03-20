# Writing the code to import the manually reported entries for vlookup during processing


import sqlite3

def create_db():
    """ create a database connection to the SQLite database
    specified by db_file
    :param db_file: database file
    :return: Connection object or None
    """
    db_file = r"fees_manual.db"

    conn = None
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)
        return
    
    sql_createtable = """ CREATE TABLE IF NOT EXISTS  (
        narr text NOT NULL,
        ref text NOT NULL,
        cr_amt text NOT NULL,
        student_name text NOT NULL); """
    

def import_csv(fname):
    pass

if __name__ = "__main__":
    create_db()