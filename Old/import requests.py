import requests
import os

en = os.environ

googleshomepage = requests.get('https://www.google.com')

print()