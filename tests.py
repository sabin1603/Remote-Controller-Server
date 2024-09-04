import os

username = os.getenv('SKYPE_EMAIL')
password = os.getenv('SKYPE_PASSWORD')

print(username)
print(password)

from skpy import Skype
sk = Skype(username, password) # connect to Skype

# print(sk.user)
contacts = sk.contacts;
chats = sk.chats

#for i in contacts:
#   print(i)

ch = sk.contacts["live:.cid.b21ef970abfde7d7"].chat # 1-to-1 conversation
ch.sendMsg("salut") # plain-text message
