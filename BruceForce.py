import itertools
import string 
from win32com.client import Dispatch
import time

file = input('Path: ')

chars = string.ascii_lowercase + string.digits + string.ascii_uppercase

for password_length in range(6, 20):
    for password in itertools.product(chars, repeat=password_length):
        password = ''.join(password)

        print ('Testing password: '+ password)
        instance = Dispatch ('Excel.Application')

        try:
            instance.Workbooks.Open(file, False, True, None, password)
            
            print ('Password Cracked: ' + password)

            break

        except:
            pass
    break      