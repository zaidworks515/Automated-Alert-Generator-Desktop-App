from db_queries import (insert_email_recipient, delete_email_recipient, search_email_recipient, get_all_recipients)

def add_email(name, email, designation):
    insert_email_recipient(name, email, designation)

def delete_email(email):
    delete_email_recipient(email)

def search_email(email):
    return search_email_recipient(email)

def list_all_emails():
    return get_all_recipients()
