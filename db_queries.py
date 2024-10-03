import sqlite3

def create_email_table():
    conn = sqlite3.connect('emails.db')
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS email_recipients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            designation TEXT NOT NULL
        )
    ''')

    conn.commit()
    conn.close()


def insert_email_recipient(name, email, designation):
    conn = sqlite3.connect('emails.db')
    cursor = conn.cursor()

    cursor.execute('''
        INSERT OR IGNORE INTO email_recipients (name, email, designation)
        VALUES (?, ?, ?)
    ''', (name, email, designation))

    conn.commit()
    conn.close()


def delete_email_recipient(email):
    conn = sqlite3.connect('emails.db')
    cursor = conn.cursor()

    cursor.execute('''
        DELETE FROM email_recipients WHERE email = ?
    ''', (email,))
    print(f"RECORD DELETED {email}")
    conn.commit()
    conn.close()


def search_email_recipient(email):
    conn = sqlite3.connect('emails.db')
    cursor = conn.cursor()

    cursor.execute('''
        SELECT * FROM email_recipients WHERE email = ?
    ''', (email,))
    
    result = cursor.fetchone()  
    
    conn.close()
    return result


def get_all_recipients():
    conn = sqlite3.connect('emails.db')
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, email, designation FROM email_recipients")  
    recipients = cursor.fetchall()
    conn.close()
    return recipients



def get_all_designations():
    conn = sqlite3.connect('emails.db')
    cursor = conn.cursor()

    cursor.execute('SELECT DISTINCT designation FROM email_recipients')
    designations = cursor.fetchall()

    conn.close()
    return [designation[0] for designation in designations]


create_email_table()
