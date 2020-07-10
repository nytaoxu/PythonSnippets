#!/usr/bin/env python
# Written by Tao

import os
import sys
import email
import imaplib
import getpass

target_titles = {
    "attachment": [],
    "keyword 1": [],
    "keyword 2": [],
}

login_credentials = {
    "username": "your email",
    "password": "your password",
    "host": "outlook.office365.com",
    "attachment_path": "attachments",
}


def is_satisfied(target_titles):
    try:
        for key in target_titles:
            if len(target_titles[key]) == 0:
                return False
        return True
    except Exception as e:
        print(f"Exception in is_satisfied(target_titles): {e}")
        return False


def includes_subject(target_titles, the_subject):
    try:
        the_subject = the_subject.lower()
        result = []
        # lower_keys = [x.lower() for x in target_titles.keys()]
        for key in target_titles.keys():
            if key.lower() in the_subject:
                result.append(key)
        return result
    except Exception as e:
        print(f"Exception in test_subject(target_titles, the_subject): {e}")
        return []


def main():
    try:
        global target_titles, login_credentials
        # print("hello world")
        # user = input("Please enter your username: ")
        username = login_credentials["username"]
        password = login_credentials["password"]
        # password = getpass.getpass("Please enter your password: ")
        # print(f"password = {password}")
        host = login_credentials["host"]
        # port = 993
        # mail = imaplib.IMAP4_SSL(host, port)
        with imaplib.IMAP4_SSL(host) as mail:
            # status, data = mail.login(*(login_credentials.values()[:2]))
            status, data = mail.login(username, password)
            # print("Success.")
            if status != "OK":
                raise Exception("Login failed.")
            # print(f"data = {data}")
            status, data = mail.select("Inbox")
            if status != "OK":
                raise Exception("Selecting 'Inbox' failed.")
            # print(f"data = {data}")
            status, data = mail.search(None, "All")
            if status != "OK":
                raise Exception("Searching emails failed.")
            # print(f"data = {data}")
            id_list = data[0].decode("utf-8").split()
            id_list.reverse()
            # id_list = id_list[:7]
            # print(f"id_list = {id_list}")
            satisfied = False
            attachment_path = os.path.join(
                ".", login_credentials["attachment_path"])
            if not os.path.exists(attachment_path):
                os.mkdir(attachment_path)
                print(
                    f"Attachment path '{attachment_path}' created successfully.")
            for email_id in id_list:
                if is_satisfied(target_titles):
                    satisfied = True
                    break
                status, data = mail.fetch(email_id, "(RFC822)")
                if status != "OK":
                    raise Exception("Fetching email failed.")
                the_email = email.message_from_bytes(data[0][1])
                # print(f"the_email = {the_email}")
                # print(dir(the_email))
                the_subject = the_email["Subject"]
                the_keys = includes_subject(target_titles, the_subject)
                if not the_keys:
                    continue
                for part in the_email.walk():
                    if part.get_content_disposition() != "attachment":
                        continue
                    file_path = os.path.join(
                        attachment_path, email_id + "_" + part.get_filename())
                    if not os.path.exists(file_path):
                        print(
                            f"Attachment '{file_path}' is being downloaded...", end="")
                        with open(file_path, "wb") as fp:
                            fp.write(part.get_payload(decode=True))
                        print(" => Success.")
                        for the_key in the_keys:
                            target_titles[the_key].append(email_id)
                # print(the_email["Subject"])
                # print(the_email["subjECT"])
                # sys.exit(0)
            if satisfied:
                print("All attachments downloaded.")
            else:
                print("Not all attachments downloaded.")
            print(f"target_titles = {target_titles}")
    except Exception as e:
        print(f"Exception in main(): {e}")


if __name__ == "__main__":
    main()
