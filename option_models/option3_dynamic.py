# pip install python-docx 
# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
from http.client import responses
import json
import pickle
from urllib import response
import pandas as pd
from docx import Document
from Google import Create_Service
from googleapiclient.errors import HttpError
import sys
import warnings
warnings.filterwarnings('ignore')

CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'forms'
API_VERSION = 'v1'
SCOPES = ["https://www.googleapis.com/auth/forms.body", "https://www.googleapis.com/auth/forms.responses.readonly"]

try:
    service = pickle.load(open('service.pkl', 'rb'))
    service
except:
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    with open('service.pkl', 'wb' ) as f:
        pickle.dump(service, f)
        f.close()

class Question:
    
    input_q = None 
    
    def __init__(self, q_id, text, type_q):
            self.q_id = q_id
            self.text = text
            self.type_q = type_q

def get_options(data, type_q, start):
    key_list = list(data.keys())[start:]
    
    options_list = []
    for key in key_list:
        
        if type_q.lower() in key.lower():
            options_list.append(data[key])
            
        else:
            break
            
    return ', '.join(options_list)

def json2df(response):
#     response = {
#         "dat": "{\"1text\":\"This is Simple Input\",\"2text\":\"This is Check Box (Multiple)\",\"3checkbox\":\"CB1\",\"4checkbox\":\"CB2\",\"5text\":\"This is DropDown (Choice)\",\"6dropdown\":\"DD1\",\"7dropdown\":\"DD2\",\"8text\":\"This is MCQ (Choice)\",\"9radio\":\"MC1\",\"10radio\":\"MC2\",\"11text\":\"This is Date\"}"
#     }

    #data = json.loads(next(iter('dat')))
    data = json.loads(response["dat"])
    
    q_dict = {}
    q_num = 1

    key_list = list(data.keys())

    for i, key in enumerate(data):

        if 'text' in key.lower():

            temp_q = Question(q_id=f'Q{q_num}', text=data[key], type_q=None)

            try:
                if 'text' in key_list[i+1].lower():
                    temp_q.type_q = 'string'

                elif 'radio' in key_list[i+1].lower():
                    temp_q.type_q = 'choice'
                    temp_q.input_q = get_options(data=data, type_q='radio', start=i+1)

                elif 'dropdown' in key_list[i+1].lower():
                    temp_q.type_q = 'choice'
                    temp_q.input_q = get_options(data=data, type_q='dropdown', start=i+1)

                elif 'checkbox' in key_list[i+1].lower():
                    temp_q.type_q = 'multiple'
                    temp_q.input_q = get_options(data=data, type_q='checkbox', start=i+1)

                else:
                    pass

            except IndexError:
                pass

            q_dict[f'q{q_num}'] = temp_q

            q_num += 1

    df = pd.DataFrame([q_dict[q].__dict__ for q in q_dict])
    
    return df

def create_question(index, title, type_q, required=True, options=None):
    
    temp_question = {
            "createItem": {
                "item": {
                    "title": title,
                    "questionItem": {
                        "question": {
                        }
                    },
                },
                "location": {
                    "index": index
                }
            }
        }

    if type_q.lower() in ['string', 'number']:
        temp_question["createItem"]['item']['questionItem']['question'] = {'required': required, 'textQuestion': {'paragraph':False}}
        
    elif type_q.lower() == 'text':
        temp_question["createItem"]['item']['questionItem']['question'] = {'required': required, 'textQuestion': {'paragraph':True}}
    
    elif type_q.lower() == 'choice':
        temp_question["createItem"]['item']['questionItem']['question'] = {'required': required, 'choiceQuestion': {'type': 'RADIO', 'options': options, 'shuffle': False}}
        
    elif type_q.lower() == 'multiple':
        temp_question["createItem"]['item']['questionItem']['question'] = {'required': required, 'choiceQuestion': {'type': 'CHECKBOX', 'options': options, 'shuffle': False}}
        
    elif type_q.lower() == 'date':
        temp_question["createItem"]['item']['questionItem']['question'] = {'required': required, 'dateQuestion': {"includeTime": False, "includeYear": True}}
    
    elif type_q.lower() == 'datetime':
        temp_question["createItem"]['item']['questionItem']['question'] = {'required': required, 'dateQuestion': {"includeTime": True, "includeYear": False}}
    
    elif type_q.lower() == 'time':
        temp_question["createItem"]['item']['questionItem']['question'] = {'required': required, 'timeQuestion': {"duration": False}}
        
    elif type_q.lower() == 'scale':
        temp_question["createItem"]['item']['questionItem']['question'] = {'required': required, 'scaleQuestion': {"high": 0, "highLabel": "High", "low": 5, "lowLabel": "Low"}}
        
    elif type_q.lower() == 'file':
        temp_question["createItem"]['item']['questionItem']['question'] = {'required': required, "fileUploadQuestion": {"maxFileSize": 1, "maxFiles": 10}}
        
    else:
        pass

    return temp_question

def create_form(response, title, service, description=None):
    
    data = json2df(response=response)
    
    NEW_FORM = {
        "info": {
            "title": title,
            "description": description
        }
    }
    
    q_list = []
    
    for i in data.index:
        if pd.notna(data.loc[i, 'input_q']):
            options = [{'value':ele.strip()} for ele in data.loc[i, 'input_q'].split(',')]
        else:
            options=None
        
        required = True if '*' in data.loc[i, 'text'] else False
        
        question = create_question(index=i, title=data.loc[i, 'text'], type_q=data.loc[i, 'type_q'], required=required, options=options)
        
        q_list.append(question)
        
    NEW_QUESTION = {"requests": q_list}
    
    # Creates the initial form
    result = service.forms().create(body=NEW_FORM).execute()
    # print('Form has been Created.')

    # Adds the question to the form
    question_setting = service.forms().batchUpdate(formId=result["formId"], body=NEW_QUESTION).execute()
    # print('Questions have been Added.')

    # Prints the result to show the question has been added
    get_result = service.forms().get(formId=result["formId"]).execute()

    # print('Visit the Responder URL:',get_result['responderUri'])
    print(str(get_result['responderUri']) + '|' + str(len(q_list)) + '|' + get_result['formId'])
    
    return get_result['formId'], get_result['responderUri']

# Usage
json_text = json.loads(json.dumps(sys.argv[1]))
file_title = sys.argv[2]

response = {
   "dat": "\"".join(json_text.split("'"))
}

# print(response)

create_form(response=response, service=service, title=file_title)

#response = {
#    "dat": "{\"1text\":\"This is Simple Input\",\"2text\":\"This is Check Box (Multiple)\",\"3checkbox\":\"CB1\",\"4checkbox\":\"CB2\",\"5text\":\"This is DropDown (Choice)\",\"6dropdown\":\"DD1\",\"7dropdown\":\"DD2\",\"8text\":\"This is MCQ (Choice)\",\"9radio\":\"MC1\",\"10radio\":\"MC2\",\"11text\":\"This is Date\"}"
#}

#create_form(response=response, service=service, title='Sample Test 1')