import requests
from pprint import pprint
from docxtpl import DocxTemplate
from flask import request


# user = 'GSAGROPAK- 17865'
# data = requests.get(
#     'http://151.80.237.86:1251/ords/zkt/exprt_doc/doc?pi_no={}'.format(user))

# data = data.json()

# # pprint(data['items'], indent=1)
# for x in data['items']:
#     pprint(x)

# print('Data is now overwriting..............')
# print(x)

import pythoncom

pythoncom.CoInitialize()
